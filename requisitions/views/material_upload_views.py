from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.utils import timezone
from ..forms import UploadFileForm, OrderModelUploadForm, MaterialDetailsUploadForm, UpdateProcessTypeDBForm, UploadInventoryFileForm
from ..models import Requisition, RequisitionItem, MaterialListVersion, WorkOrderMaterial, Inventory, MachineModel, ProcessType
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

@login_required
def upload_materials(request, pk):
    requisition = get_object_or_404(Requisition, pk=pk)

    is_admin = request.user.is_superuser
    is_material_handler = request.user.groups.filter(name='撥料人員').exists()

    if not is_material_handler and not is_admin:
        messages.error(request, "您沒有權限上傳物料清單。")
        return redirect('requisition_list')

    if request.method == 'POST':
        form = UploadFileForm(request.POST, request.FILES)
        if form.is_valid():
            excel_file = request.FILES['file']
            try:
                with transaction.atomic():
                    new_material_version = MaterialListVersion.objects.create(
                        requisition=requisition,
                        uploaded_by=request.user,
                    )

                    requisition.current_material_list_version = new_material_version
                    if requisition.status != 'pending':
                        requisition.status = 'pending'
                        requisition.material_confirmed_by = None
                        requisition.material_confirmed_date = None
                        requisition.sign_off_by = None
                        requisition.sign_off_date = None
                    requisition.save()

                    df = pd.read_excel(excel_file)
                    df.columns = df.columns.str.strip()

                    column_map = {
                        '訂單': 'order_number',
                        '訂單單號': 'order_number',
                        '物料': 'material_number',
                        '品名': 'item_name',
                        '物料說明': 'item_name',
                        '機型': 'machine_model',
                        '需求數量': 'required_quantity',
                        '庫存數量': 'stock_quantity',
                    }
                    df.rename(columns=column_map, inplace=True)

                    required_cols = ['order_number', 'material_number', 'item_name', 'machine_model', 'required_quantity']
                    missing_cols = [col for col in required_cols if col not in df.columns]
                    if missing_cols:
                        # Map back to original names for error message
                        original_missing = [key for key, val in column_map.items() if val in missing_cols]
                        raise ValueError(f"上傳的 Excel 檔案中缺少必要的欄位，請檢查是否包含： {', '.join(original_missing)}")

                    df_aggregated = df.groupby([
                        'order_number', 'material_number', 'machine_model'
                    ]).agg({
                        'required_quantity': 'sum',
                        'item_name': 'first',
                        'stock_quantity': 'first'
                    }).reset_index()

                    # Create a list of tuples for all materials to fetch
                    materials_to_find = [
                        (str(row['order_number']).strip(), str(row['material_number']).strip(), str(row['machine_model']).strip())
                        for index, row in df_aggregated.iterrows()
                    ]

                    found_materials = []
                    batch_size = 500
                    for i in range(0, len(materials_to_find), batch_size):
                        batch = materials_to_find[i:i + batch_size]
                        query = Q()
                        for order, material, machine in batch:
                            query |= Q(order_number=order, material_number=material, machine_model__name=machine)
                        found_materials.extend(WorkOrderMaterial.objects.filter(query).select_related('machine_model'))

                    # Create a dictionary for quick lookup
                    materials_dict = {
                        (m.order_number, m.material_number, m.machine_model.name): m
                        for m in found_materials
                    }

                    items_to_create = []
                    not_found_materials = []

                    for index, row in df_aggregated.iterrows():
                        order_number = str(row['order_number']).strip()
                        material_number = str(row['material_number']).strip()
                        machine_model_name = str(row['machine_model']).strip()

                        work_order_material = materials_dict.get((order_number, material_number, machine_model_name))

                        if work_order_material:
                            items_to_create.append(
                                RequisitionItem(
                                    material_list_version=new_material_version,
                                    source_material=work_order_material,
                                    order_number=order_number,
                                    material_number=material_number,
                                    item_name=row['item_name'],
                                    required_quantity=row['required_quantity'],
                                    stock_quantity=row.get('stock_quantity', 0),
                                    confirmed_quantity=None,
                                    is_signed_off=False,
                                )
                            )
                        else:
                            not_found_materials.append(f"訂單 {order_number}, 物料 {material_number}, 機型 {machine_model_name}")

                    if not_found_materials:
                        error_message = "上傳失敗，因為在主物料清單中找不到以下物料，請檢查資料是否正確: <br><ul>" + "".join([f"<li>{item}</li>" for item in not_found_materials]) + "</ul>"
                        messages.error(request, error_message, extra_tags='safe')
                        raise IntegrityError("Aborting transaction due to missing source materials.")

                    RequisitionItem.objects.bulk_create(items_to_create)

                messages.success(request, "物料清單上傳成功！")
                return redirect('requisition_list')

            except ValueError as ve:
                messages.error(request, str(ve))
            except IntegrityError:
                pass
            except Exception as e:
                messages.error(request, f"上傳檔案時發生未預期的錯誤: {e}")
    else:
        form = UploadFileForm()
    
    material_versions = requisition.material_versions.all().order_by('-uploaded_at')

    return render(request, 'requisitions/upload_materials.html', {
        'form': form, 
        'requisition': requisition,
        'material_versions': material_versions,
    })


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
              demand_date_col = form.cleaned_data['demand_date_col'] # Get demand_date_col from form

              try:
                  # Step 1: Read the process type mapping from the local DB file
                  try:
                      db_path = os.path.join(settings.BASE_DIR, 'output.xlsx')
                      excel_sheets = pd.read_excel(db_path, engine='openpyxl',
  sheet_name=None)
                      df_db = pd.concat(excel_sheets.values(), ignore_index=True)

                      # Ensure required columns exist in output.xlsx
                      if '物料' not in df_db.columns or '機型' not in df_db.columns or '投料點' not in df_db.columns:
                          raise ValueError("output.xlsx 檔案中必須包含 '物料', '機型','投料點' 欄位。")

                      df_db['material_prefix'] = df_db['物料'].astype(str).str[:10]
                      df_db['machine_model_name'] = df_db['機型'].astype(str).str.strip()

                      # Create a composite key for lookup
                      df_db['composite_key'] = list(zip(df_db['material_prefix'],
  df_db['machine_model_name']))

                      # Create the mapping: (material_prefix, machine_model_name) ->process_type_name
                      process_type_map =df_db.set_index('composite_key')['投料點'].to_dict()

                  except Exception as e:
                      messages.error(request, f"讀取投料點資料庫 (output.xlsx) 時發生錯誤:{e}")
                      return redirect('requisitions:upload_material_details_excel') # Redirect to the specific view

                  # Step 2: Read the uploaded Excel file
                  df_upload = pd.read_excel(excel_file, dtype=str, engine='openpyxl')
                  df_upload.columns = df_upload.columns.str.strip()

                  # Step 3: Validate required columns
                  order_col = '訂單單號' if '訂單單號' in df_upload.columns else '訂單'
                  if order_col not in df_upload.columns:
                      raise ValueError("上傳的 Excel 檔案中找不到 '訂單單號' 或 '訂單'欄位。")
                  if '物料' not in df_upload.columns:
                      raise ValueError("上傳的 Excel 檔案中找不到 '物料' 欄位。")
                  if required_qty_col not in df_upload.columns: 
                      raise ValueError(f"在 Excel 中找不到您指定的 '需求數量'欄位：'{required_qty_col}'。")
                  # New validation for demand_date_col
                  if demand_date_col and demand_date_col not in df_upload.columns:
                      raise ValueError(f"在 Excel 中找不到您指定的 '需求日期'欄位：'{demand_date_col}'。")

                  df_upload[required_qty_col] = pd.to_numeric(df_upload[required_qty_col], errors='coerce').fillna(0)

                  # Group by order and material, summing the required quantity
                  df_aggregated = df_upload.groupby([order_col, '物料']).agg({
                      required_qty_col: 'sum',
                      '物料說明': 'first',  # Keep the first item name found
                      demand_date_col: 'first', # Include demand_date_col in aggregation
                  }).reset_index()

                  created_count = 0
                  updated_count = 0

                  with transaction.atomic():
                      order_numbers_in_upload = df_aggregated[order_col].astype(str).str.strip().unique()
                      
                      # Mark existing materials for these orders as inactive. We will reactivate or update them.
                      WorkOrderMaterial.objects.filter(order_number__in=order_numbers_in_upload).update(is_active=False)

                      # Process each aggregated row
                      for _, row in df_aggregated.iterrows():
                          order_number_clean = str(row.get(order_col)).strip()
                          material_number_clean = str(row.get('物料')).strip()

                          if not all([order_number_clean, material_number_clean]):
                              continue

                          parent_scope_entry = WorkOrderMaterial.objects.filter(
                              order_number=order_number_clean,
                              material_number="PARENT_SCOPE"
                          ).first()

                          if not parent_scope_entry or not parent_scope_entry.machine_model:
                              raise ValueError(f"訂單 {order_number_clean} 的父階範圍不存在或缺少機型資訊。請先上傳訂單與機型 Excel。")

                          machine_model_obj = parent_scope_entry.machine_model
                          machine_model_name_clean = machine_model_obj.name

                          material_prefix = material_number_clean[:10]
                          composite_lookup_key = (material_prefix, machine_model_name_clean)
                          process_type_name = str(process_type_map.get(composite_lookup_key, '其他')).strip()

                          process_type_obj, _ = ProcessType.objects.get_or_create(
                              name=process_type_name,
                              machine_model=machine_model_obj
                          )

                          item_name_clean = str(row.get('物料說明', '')).strip()
                          required_quantity_clean = row.get(required_qty_col, 0)

                          # Parse demand_date
                          demand_date_clean = None
                          if demand_date_col:
                              demand_date_str = row.get(demand_date_col)
                              if demand_date_str:
                                  try:
                                      # Attempt to parse various date formats
                                      demand_date_clean = pd.to_datetime(demand_date_str).date()
                                  except ValueError:
                                      messages.warning(request, f"訂單 {order_number_clean}, 物料 {material_number_clean}: 無效的需求日期格式 '{demand_date_str}'，將跳過此日期。")
                                  except Exception as e:
                                      messages.warning(request, f"訂單 {order_number_clean}, 物料 {material_number_clean}: 解析需求日期時發生錯誤 '{demand_date_str}': {e}，將跳過此日期。")

                          # Custom logic to handle potential duplicates and merge them
                          existing_materials = WorkOrderMaterial.objects.filter(
                              order_number=order_number_clean,
                              material_number=material_number_clean,
                              machine_model=machine_model_obj,
                              process_type=process_type_obj
                          )

                          if existing_materials.exists():
                              # Merge duplicates if they exist
                              master_record = existing_materials.first()
                              other_records = existing_materials.exclude(pk=master_record.pk)

                              total_confirmed = master_record.confirmed_quantity or 0
                              for record in other_records:
                                  total_confirmed += record.confirmed_quantity or 0
                                  record.transactions.update(work_order_material=master_record)

                              # Update the master record
                              master_record.item_name = item_name_clean
                              master_record.required_quantity = required_quantity_clean
                              master_record.confirmed_quantity = total_confirmed
                              master_record.is_active = True
                              master_record.demand_date = demand_date_clean # Assign demand_date
                              master_record.save()

                              # Delete the now-redundant records
                              other_records.delete()
                              updated_count += 1
                          else:
                              # No existing material, create a new one
                              WorkOrderMaterial.objects.create(
                                  order_number=order_number_clean,
                                  material_number=material_number_clean,
                                  machine_model=machine_model_obj,
                                  process_type=process_type_obj,
                                  item_name=item_name_clean,
                                  required_quantity=required_quantity_clean,
                                  is_active=True,
                                  demand_date=demand_date_clean # Assign demand_date
                              )
                              created_count += 1

                  messages.success(request, f"物料明細同步成功！新增 {created_count} 筆，更新 {updated_count} 筆物料。")
                  return redirect('core:homepage')

              except Exception as e:
                  messages.error(request, f"上傳檔案時發生錯誤: {e}")
                  import traceback
                  print(traceback.format_exc())

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