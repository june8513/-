from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.contrib.auth import authenticate, login, logout
from django.contrib import messages
from django.utils import timezone
from .forms import RequisitionForm, UploadFileForm, OrderModelUploadForm, MaterialDetailsUploadForm, RequisitionItemMaterialConfirmationFormSet, RequisitionItemSignOffFormSet, UpdateProcessTypeDBForm, UploadInventoryFileForm, ProcessTypeForm, RequisitionImageUploadForm, WorkOrderMaterialImageUploadForm, UploadStorageBinFileForm
from .models import Requisition, RequisitionItem, MaterialListVersion, WorkOrderMaterial, Inventory, MachineModel, ProcessType, RequisitionImage, WorkOrderMaterialTransaction, WorkOrderMaterialImage
from django.db import transaction
import openpyxl
import os
from django.conf import settings
from django.contrib.auth.forms import AuthenticationForm
from django.contrib.auth.models import User, Group
from django.db.models import Q, F, Value, DecimalField, OuterRef, Subquery, Exists, ExpressionWrapper
from django.db.models.functions import Coalesce
from django.core.paginator import Paginator, EmptyPage, PageNotAnInteger # Import Paginator
from django.urls import reverse
from django.db import IntegrityError
import pandas as pd
from django.http import JsonResponse, HttpResponse
from django.template.loader import render_to_string
import io
import json
from decimal import Decimal
import tempfile # Import tempfile
from requisitions.utils import process_order_model_excel, process_material_details_excel # Import the utility functions


@login_required
def finished_goods_dispatch(request):
    return render(request, 'requisitions/finished_goods_dispatch.html')

@login_required
def view_requisition_images(request, pk):
    requisition = get_object_or_404(Requisition, pk=pk)
    images = requisition.images.all()
    return render(request, 'requisitions/view_requisition_images.html', {'requisition': requisition, 'images': images})

@login_required
def upload_requisition_images_page(request, pk):
    requisition = get_object_or_404(Requisition, pk=pk)
    form = RequisitionImageUploadForm()
    return render(request, 'requisitions/upload_requisition_images.html', {'requisition': requisition, 'form': form})


@login_required
def view_process_type_database(request):
    if not request.user.is_superuser:
        messages.error(request, "您沒有權限執行此操作。")
        return redirect('homepage')
    db_path = os.path.join(settings.BASE_DIR, 'output.xlsx')
    data = []
    headers = []
    try:
        excel_sheets = pd.read_excel(db_path, engine='openpyxl', sheet_name=None)
        df_db = pd.concat(excel_sheets.values(), ignore_index=True)
        headers = df_db.columns.tolist()
        data = df_db.to_dict(orient='records')
    except FileNotFoundError:
        messages.error(request, "投料點資料庫檔案 (output.xlsx) 不存在。")
    except Exception as e:
        messages.error(request, f"讀取投料點資料庫時發生錯誤: {e}")
    context = {
        'data': data,
        'headers': headers,
    }
    return render(request, 'requisitions/view_process_type_database.html', context)



@login_required
def clear_work_order_material_database(request):
    if not request.user.is_superuser:
        messages.error(request, "您沒有權限執行此操作。")
        return redirect('homepage')
    if request.method == 'POST':
        WorkOrderMaterial.objects.all().delete()
        messages.success(request, "訂單主物料清單資料庫已成功清空。")
    return redirect('view_database')



@login_required
def view_database(request):
    if not request.user.is_superuser:
        messages.error(request, "您沒有權限執行此操作。")
        return redirect('homepage')
    materials = WorkOrderMaterial.objects.all().select_related('process_type', 'machine_model').select_related('process_type', 'machine_model').select_related('process_type', 'machine_model').select_related('process_type', 'machine_model').select_related('process_type', 'machine_model')
    context = {
        'materials': materials,
    }
    return render(request, 'requisitions/view_database.html', context)



def _filter_requisitions(request, sort_by='created_at', order='desc', material_status_filter=None):
    """
    Helper function to filter requisitions based on user role and query parameters.
    NOW FILTERS FOR UNDISPATCHED REQUISITIONS ONLY.
    """
    is_admin = request.user.is_superuser
    is_applicant = request.user.groups.filter(name='申請人員').exists()
    is_material_handler = request.user.groups.filter(name='撥料人員').exists()

    order_field = sort_by
    if order == 'desc':
        order_field = '-' + order_field

    # CORE CHANGE: Only show requisitions that have NOT been dispatched.
    base_queryset = Requisition.objects.filter(dispatch_performed=False).order_by(order_field).select_related('applicant')

    # Keep role-based filtering, but on the undispatched list
    if is_admin:
        all_requisitions = base_queryset
    elif is_applicant:
        # Applicants see their own undispatched requisitions
        all_requisitions = base_queryset.filter(applicant=request.user)
    elif is_material_handler:
        # Material handlers see all undispatched requisitions
        all_requisitions = base_queryset
    else:
        all_requisitions = Requisition.objects.none()

    process_type_filter = request.GET.get('process_type')
    if process_type_filter:
        all_requisitions = all_requisitions.filter(process_type=process_type_filter)

    # The status filter is now redundant as we use dispatch_performed
    # We can keep the material_status_filter for more granular control on the pending page
    if material_status_filter:
        if material_status_filter == 'has_materials':
            all_requisitions = all_requisitions.filter(
                current_material_list_version__isnull=False,
                current_material_list_version__items__isnull=False
            ).distinct()
        elif material_status_filter == 'no_materials':
            all_requisitions = all_requisitions.filter(
                Q(current_material_list_version__isnull=True) |
                Q(current_material_list_version__items__isnull=True)
            ).distinct()

    unique_requisitions = []
    seen_combinations = set()
    for req in all_requisitions: # all_requisitions is the filtered queryset
        combination = (req.order_number, req.process_type)
        if combination not in seen_combinations:
            # Fetch unique machine models for this requisition based on order number
            machine_model_names = list(WorkOrderMaterial.objects.filter(
                order_number=req.order_number
            ).values_list('machine_model__name', flat=True).distinct().order_by('machine_model__name'))
            
            req.machine_models_display = ", ".join(machine_model_names)
            unique_requisitions.append(req)
            seen_combinations.add(combination)
    
    return unique_requisitions

@login_required
def requisition_list(request):
    sort_by = request.GET.get('sort_by', 'process_type') # Default sort by process_type
    order = request.GET.get('order', 'asc') # Default order ascending for process_type

    # Map frontend sort_by names to model field names
    sort_mapping = {
        'work_order_number': 'order_number',
        'applicant': 'applicant__username',
        'request_date': 'request_date',
        'process_type': 'process_type',
        'status': 'status',
        'created_at': 'created_at',
    }
    model_sort_by = sort_mapping.get(sort_by, 'process_type') # Default to process_type if invalid sort_by

    process_type_selected = request.GET.get('process_type')
    status_selected = request.GET.get('status', 'pending') # Default status to 'pending'
    material_status_selected = request.GET.get('material_status')

    # show_results should always be true, the template will handle empty results
    show_results = True 
    
    unique_requisitions = _filter_requisitions(request, 
                                                sort_by=model_sort_by, 
                                                order=order, 
                                                material_status_filter=material_status_selected)

    paginator = Paginator(unique_requisitions, 10)
    page = request.GET.get('page')
    try:
        requisitions_page = paginator.page(page)
    except PageNotAnInteger:
        requisitions_page = paginator.page(1)
    except EmptyPage:
        requisitions_page = paginator.page(paginator.num_pages)

    material_status_choices = [
        ('', '所有待撥料'),
        ('has_materials', '已上傳物料'),
        ('no_materials', '未上傳物料'),
    ]

    process_types = Requisition.objects.order_by('process_type').values_list('process_type', flat=True).distinct()
    process_type_choices = [(pt, pt) for pt in process_types if pt]

    context = {
        'requisitions': requisitions_page,
        'is_admin': request.user.is_superuser,
        'is_applicant': request.user.groups.filter(name='申請人員').exists(),
        'is_material_handler': request.user.groups.filter(name='撥料人員').exists(),
        'status_choices': Requisition.STATUS_CHOICES,
        'process_type_choices': process_type_choices,
        'material_status_choices': material_status_choices, # New choices for template
        'selected_status': status_selected,
        'selected_process_type': process_type_selected,
        'selected_material_status': material_status_selected, # Pass selected material status
        'sort_by': sort_by,
        'order': order,
        'query_params': request.GET.urlencode(),
        'show_results': show_results,
    }
    return render(request, 'requisitions/requisition_list.html', context)

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
def requisition_create(request):
    if not request.user.groups.filter(name='申請人員').exists() and not request.user.is_superuser:
        messages.error(request, "您沒有權限建立撥料申請單。")
        return redirect('requisition_list')

    if request.method == 'POST':
        order_number = request.POST.get('order_number') # Get order_number from POST data

        # Re-generate choices for process_type based on the submitted order_number
        material_process_type_ids = WorkOrderMaterial.objects.filter(
            order_number=order_number
        ).values_list('process_type__id', flat=True).distinct()

        used_requisition_process_type_names = Requisition.objects.filter(
            order_number=order_number
        ).values_list('process_type', flat=True)

        available_process_types_query = ProcessType.objects.filter(
            id__in=material_process_type_ids
        ).exclude(
            name__in=used_requisition_process_type_names
        ).order_by('name')
        
        # Format choices for Django form: [(value, label), ...]
        form_process_type_choices = [(pt.id, pt.name) for pt in available_process_types_query]

        form = RequisitionForm(request.POST, process_type_choices=form_process_type_choices)
        if form.is_valid():
            try:
                # order_number is already defined
                # process_type is now handled by the form's cleaned_data
                
                existing_requisition = Requisition.objects.filter(
                    order_number=order_number,
                    process_type=form.cleaned_data['process_type'] # Use cleaned_data here
                ).first()

                if existing_requisition:
                    messages.error(request, "此訂單單號在該需求流程中已存在，請選擇不同的訂單單號或需求流程，或修改現有申請單。")
                    return render(request, 'requisitions/requisition_create.html', {'form': form})

                # Get the ProcessType object using the ID from the form
                selected_process_type_id = form.cleaned_data['process_type']
                selected_process_type_obj = get_object_or_404(ProcessType, id=selected_process_type_id)

                requisition = form.save(commit=False)
                requisition.applicant = request.user
                requisition.order_number = order_number
                requisition.process_type = selected_process_type_obj.name # Assign the name to the CharField
                requisition.save()
                messages.success(request, "撥料申請單建立成功！")
                return redirect('requisition_list')
            except IntegrityError:
                messages.error(request, "此訂單單號在該需求流程中已存在，請使用不同的訂單單號或需求流程。")
            except Exception as e:
                messages.error(request, f"建立撥料申請單時發生錯誤: {e}")
                import traceback
                print(traceback.format_exc())
    else:
        form = RequisitionForm()
    return render(request, 'requisitions/requisition_create.html', {'form': form})


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
        return redirect('homepage')

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
                return redirect('homepage')

            except Exception as e:
                messages.error(request, f"上傳檔案時發生錯誤: {e}")
                import traceback
                print(traceback.format_exc())
            finally:
                # Clean up the temporary file
                os.unlink(temp_file_path)

    else:
        form = OrderModelUploadForm()
    
    return render(request, 'requisitions/upload_order_model_excel.html', {'form': form})


@login_required
def upload_material_details_excel(request):
      if not request.user.is_superuser:
          messages.error(request, "您沒有權限執行此操作。")
          return redirect('homepage')

      if request.method == 'POST':
          form = MaterialDetailsUploadForm(request.POST, request.FILES)
          if form.is_valid():
              excel_file = request.FILES['file']
              required_qty_col = form.cleaned_data['required_quantity_col']

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
                      return redirect('upload_material_details_excel')

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

                  df_upload[required_qty_col] = pd.to_numeric(df_upload[required_qty_col], errors='coerce').fillna(0)

                  # Group by order and material, summing the required quantity
                  df_aggregated = df_upload.groupby([order_col, '物料']).agg({
                      required_qty_col: 'sum',
                      '物料說明': 'first'  # Keep the first item name found
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
                                  is_active=True
                              )
                              created_count += 1

                  messages.success(request, f"物料明細同步成功！新增 {created_count} 筆，更新 {updated_count} 筆物料。")
                  return redirect('homepage')

              except Exception as e:
                  messages.error(request, f"上傳檔案時發生錯誤: {e}")
                  import traceback
                  print(traceback.format_exc())

      else:
          form = MaterialDetailsUploadForm()

      context = {'form': form}
      return render(request, 'requisitions/upload_material_details_excel.html', context)



@login_required
def material_confirmation(request, pk):
    requisition = get_object_or_404(Requisition, pk=pk)

    is_admin = request.user.is_superuser
    is_material_handler = request.user.groups.filter(name='撥料人員').exists()

    if not is_material_handler and not is_admin:
        messages.error(request, "您沒有權限執行物料確認操作。")
        return redirect('requisition_list')

    if requisition.status != 'pending':
        messages.warning(request, f"此申請單狀態為 '{requisition.get_status_display()}'，無法進行物料確認。")
        return redirect('requisition_list')

    # Sorting logic
    sort_by = request.GET.get('sort_by', 'material_number')
    order = request.GET.get('order', 'asc')
    sort_mapping = {
        'material_number': 'material_number',
        'item_name': 'item_name',
        'required_quantity': 'required_quantity',
    }
    model_sort_by = sort_mapping.get(sort_by, 'material_number')
    order_field = f'{'-' if order == 'desc' else ''}{model_sort_by}'

    queryset = RequisitionItem.objects.filter(
        material_list_version=requisition.current_material_list_version
    ).order_by(order_field)
    print("Queryset count:", queryset.count()) # Debugging
    formset = RequisitionItemMaterialConfirmationFormSet(queryset=queryset)
    print("Formset is bound:", formset.is_bound) # Debugging
    print("Formset total forms:", formset.total_form_count()) # Debugging
    
    # Fetch inventory data for each item and attach to form
    for form in formset:
        try:
            inventory_item = Inventory.objects.get(material_number=form.instance.material_number)
            form.inventory_item = inventory_item # Attach inventory item to the form object
        except Inventory.DoesNotExist:
            form.inventory_item = None # Set to None if not found

    # Get unique machine models for this requisition
    unique_machine_model_names = []
    if requisition.current_material_list_version:
        # Get machine models from the source_material of RequisitionItems
        machine_model_ids = RequisitionItem.objects.filter(
            material_list_version=requisition.current_material_list_version,
            source_material__machine_model__isnull=False # Ensure there's a machine model
        ).values_list('source_material__machine_model__id', flat=True).distinct()

        unique_machine_models = MachineModel.objects.filter(id__in=machine_model_ids).order_by('name')
        unique_machine_model_names = [str(mm.name) for mm in unique_machine_models] # Get names

    if request.method == 'POST':
        formset = RequisitionItemMaterialConfirmationFormSet(request.POST, queryset=queryset)
        print("Request POST data:", request.POST) # Add this line for debugging
        if formset.is_valid():
            items = formset.save()
            
            for item in items:
                if item.confirmed_quantity is not None and item.confirmed_quantity > item.required_quantity:
                    messages.warning(request, f"物料 {item.material_number} 的撥料數量 ({item.confirmed_quantity}) 超過需求數量 ({item.required_quantity})。")
                if item.source_material and item.confirmed_quantity is not None:
                    item.source_material.confirmed_quantity = item.confirmed_quantity
                    item.source_material.save()

            all_items_confirmed = all(item.confirmed_quantity is not None for item in queryset)
            
            if all_items_confirmed:
                with transaction.atomic():
                    requisition.status = 'materials_confirmed'
                    requisition.material_confirmed_by = request.user
                    requisition.material_confirmed_date = timezone.now()
                    requisition.save()
                    messages.success(request, "物料已全部確認，申請單狀態已更新！")
                return redirect('requisition_list')
            else:
                messages.info(request, "物料確認已保存，但仍有未確認項目。")
                return redirect('material_confirmation', pk=requisition.pk)
        else:
            print("Formset errors:", formset.errors)
            print("Formset non-form errors:", formset.non_form_errors) # Add this line
            print("Formset management form errors:", formset.management_form.errors) # Add this line
            messages.error(request, "物料確認保存失敗，請檢查輸入。")
    
    return render(request, 'requisitions/material_confirmation.html', {
        'requisition': requisition,
        'formset': formset,
        'unique_machine_model_names': unique_machine_model_names, # Pass to context
    })

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
    output.seek(0)

    response = HttpResponse(
        output,
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = f'attachment; filename="material_confirmation_{requisition.order_number}.xlsx"'
    return response


@login_required
def requisition_sign_off(request, pk, version_pk=None):
    requisition = get_object_or_404(Requisition, pk=pk)

    is_admin = request.user.is_superuser
    is_applicant = request.user.groups.filter(name='申請人員').exists()

    if not is_applicant and not is_admin:
        messages.error(request, "您沒有權限執行最終簽收操作。")
        return redirect('requisition_list')

    if version_pk:
        target_version = get_object_or_404(MaterialListVersion, pk=version_pk, requisition=requisition)
    else:
        target_version = requisition.material_versions.order_by('-uploaded_at').first()

    if not target_version:
        messages.error(request, "找不到要簽收的物料清單版本。")
        return redirect('requisition_detail', pk=requisition.pk)

    target_version_items = RequisitionItem.objects.filter(material_list_version=target_version)
    
    # Fetch inventory data for each item and attach to form
    formset = RequisitionItemSignOffFormSet(queryset=target_version_items)
    for form in formset:
        try:
            inventory_item = Inventory.objects.get(material_number=form.instance.material_number)
            form.inventory_item = inventory_item # Attach inventory item to the form object
        except Inventory.DoesNotExist:
            form.inventory_item = None # Set to None if not found

    all_items_confirmed_in_target_version = all(item.confirmed_quantity is not None for item in target_version_items)

    if not all_items_confirmed_in_target_version:
        messages.warning(request, "此物料清單版本中的所有物料尚未確認，無法進行最終簽收。")
        return redirect('requisition_detail', pk=requisition.pk)

    queryset = target_version_items

    if request.method == 'POST':
        formset = RequisitionItemSignOffFormSet(request.POST, queryset=queryset)
        if formset.is_valid():
            items = formset.save()

            for item in items:
                if item.source_material:
                    if item.is_signed_off:
                        item.source_material.is_signed_off = True
                        if item.confirmed_quantity is not None:
                            item.source_material.confirmed_quantity = item.confirmed_quantity
                    item.source_material.save()

            all_items_signed_off = all(item.is_signed_off for item in queryset)

            if all_items_signed_off:
                with transaction.atomic():
                    if target_version == requisition.current_material_list_version:
                        requisition.status = 'completed'
                        requisition.sign_off_by = request.user
                        requisition.sign_off_date = timezone.now()
                        requisition.save()
                        messages.success(request, "撥料申請單已全部最終簽收！")
                    else:
                        messages.info(request, f"物料清單版本 '{target_version.uploaded_at.strftime('%Y-%m-%d %H:%M')}' 已全部最終簽收。")

                return redirect('requisition_detail', pk=requisition.pk)
            else:
                messages.info(request, "最終簽收已保存，但仍有未簽收項目。")
                if version_pk:
                    return redirect('requisition_sign_off_version', pk=requisition.pk, version_pk=version_pk)
                else:
                    return redirect('requisition_sign_off', pk=requisition.pk)
        else:
            messages.error(request, "最終簽收保存失敗，請檢查輸入。")
    
    return render(request, 'requisitions/requisition_sign_off.html', {'requisition': requisition, 'formset': formset, 'target_version': target_version})


@login_required
def requisition_history(request):
    is_admin = request.user.is_superuser
    is_applicant = request.user.groups.filter(name='申請人員').exists()
    is_material_handler = request.user.groups.filter(name='撥料人員').exists()

    if not is_admin and not is_applicant and not is_material_handler:
        messages.error(request, "您沒有權限查看此頁面。")
        return redirect('homepage')

    history_requisitions_qs = Requisition.objects.filter(dispatch_performed=True).select_related('applicant').order_by('-updated_at')

    work_order_number = request.GET.get('work_order_number')
    applicant_username = request.GET.get('applicant_username')
    process_type = request.GET.get('process_type')
    start_date = request.GET.get('start_date')
    end_date = request.GET.get('end_date')
    material_or_item_search = request.GET.get('material_or_item_search')

    if work_order_number:
        history_requisitions_qs = history_requisitions_qs.filter(order_number__icontains=work_order_number)
    if applicant_username:
        history_requisitions_qs = history_requisitions_qs.filter(applicant__username__icontains=applicant_username)
    if process_type:
        history_requisitions_qs = history_requisitions_qs.filter(process_type=process_type)
    if start_date:
        history_requisitions_qs = history_requisitions_qs.filter(request_date__gte=start_date)
    if end_date:
        history_requisitions_qs = history_requisitions_qs.filter(request_date__lte=end_date)
    
    if material_or_item_search:
        latest_material_version_subquery = Subquery(
            MaterialListVersion.objects.filter(
                requisition=OuterRef('pk')
            ).order_by('-uploaded_at').values('pk')[:1]
        )
        matching_items_subquery = RequisitionItem.objects.filter(
            Q(material_number__icontains=material_or_item_search) |
            Q(item_name__icontains=material_or_item_search),
            material_list_version__pk=latest_material_version_subquery)
        history_requisitions_qs = history_requisitions_qs.filter(
            Exists(matching_items_subquery.filter(material_list_version__requisition=OuterRef('pk')))
        ).distinct()

    paginator = Paginator(history_requisitions_qs, 10)
    page_number = request.GET.get('page')
    requisitions_page = paginator.get_page(page_number)

    for req in requisitions_page:
        machine_model_names = list(WorkOrderMaterial.objects.filter(
            order_number=req.order_number
        ).values_list('machine_model__name', flat=True).distinct().order_by('machine_model__name'))
        req.machine_models_display = ", ".join(machine_model_names)

    return render(request, 'requisitions/requisition_history.html', {
        'history_requisitions': requisitions_page,
    })


def user_login(request):
    if request.method == 'POST':
        form = AuthenticationForm(request, data=request.POST)
        if form.is_valid():
            username = form.cleaned_data.get('username')
            password = form.cleaned_data.get('password')
            user = authenticate(username=username, password=password)
            if user is not None:
                login(request, user)
                messages.success(request, f"歡迎回來, {username}!")
                return redirect('homepage')
            else:
                messages.error(request, "無效的使用者名稱或密碼。")
        else:
            messages.error(request, "無效的使用者名稱或密碼。")
    else:
        form = AuthenticationForm()
    return render(request, 'requisitions/login.html', {'form': form})


@login_required
def user_logout(request):
    logout(request)
    messages.info(request, "您已成功登出。")
    return redirect('login')



@login_required
def activate_material_version(request, pk, version_pk):
    requisition = get_object_or_404(Requisition, pk=pk)

    is_admin = request.user.is_superuser
    is_material_handler = request.user.groups.filter(name='撥料人員').exists()

    if not is_material_handler and not is_admin:
        messages.error(request, "您沒有權限激活物料清單版本。")
        return redirect('requisition_detail', pk=requisition.pk)

    old_version = get_object_or_404(MaterialListVersion, pk=version_pk, requisition=requisition)

    if request.method == 'POST':
        try:
            with transaction.atomic():
                new_material_version = MaterialListVersion.objects.create(
                    requisition=requisition,
                    uploaded_by=request.user,
                )

                for item in old_version.items.all():
                    RequisitionItem.objects.create(
                        material_list_version=new_material_version,
                        source_material=item.source_material, # Also copy the source material link
                        order_number=item.order_number,
                        material_number=item.material_number,
                        item_name=item.item_name,
                        required_quantity=item.required_quantity,
                        stock_quantity=item.stock_quantity,
                        confirmed_quantity=None, # Always reset confirmed quantity
                        is_signed_off=False, # Always reset sign-off status
                    )
                
                requisition.current_material_list_version = new_material_version
                requisition.status = 'pending'
                requisition.material_confirmed_by = None
                requisition.material_confirmed_date = None
                requisition.sign_off_by = None
                requisition.sign_off_date = None
                requisition.save()

                messages.success(request, "物料清單版本已成功激活，並已重置申請單狀態為待撥料。")
                return redirect('material_confirmation', pk=requisition.pk)
        except Exception as e:
            messages.error(request, f"激活物料清單版本時發生錯誤: {e}")
    
    return redirect('requisition_detail', pk=requisition.pk)


def get_available_process_types(request):
    order_number = request.GET.get('order_number')
    if not order_number:
        return JsonResponse({'error': 'No order number provided'}, status=400)

    # Get process types associated with materials for this order number
    material_process_type_ids = WorkOrderMaterial.objects.filter(
        order_number=order_number
    ).values_list('process_type__id', flat=True).distinct()

    # Get process types already used for this order number in existing requisitions
    used_requisition_process_type_names = Requisition.objects.filter(
        order_number=order_number
    ).values_list('process_type', flat=True) # This stores the name of the process type

    # Get all available process types from the database that are linked to materials for this order
    # and are not already used in existing requisitions for this order
    available_process_types_query = ProcessType.objects.filter(
        id__in=material_process_type_ids
    ).exclude(
        name__in=used_requisition_process_type_names
    ).order_by('name')
    
    available_process_types_list = []
    seen_names = set()
    for pt in available_process_types_query:
        if pt.name not in seen_names:
            available_process_types_list.append({'value': pt.id, 'label': pt.name})
            seen_names.add(pt.name)
    
    return JsonResponse({'available_process_types': available_process_types_list})


def homepage(request):
    if request.user.is_authenticated:
        is_admin = request.user.is_superuser
        is_applicant = request.user.groups.filter(name='申請人員').exists()
        is_material_handler = request.user.groups.filter(name='撥料人員').exists()
        
        context = {
            'is_admin': is_admin,
            'is_applicant': is_applicant,
            'is_material_handler': is_material_handler,
        }
        return render(request, 'requisitions/homepage.html', context)
    else:
        return render(request, 'requisitions/landing.html')


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


@login_required
def requisition_delete(request, pk):
    requisition = get_object_or_404(Requisition, pk=pk)
    is_admin = request.user.is_superuser

    if not is_admin:
        messages.error(request, "您沒有權限刪除撥料申請單。")
        return redirect('requisition_list')

    if request.method == 'POST':
        requisition.delete()
        messages.success(request, "撥料申請單已成功刪除。")
        return redirect('requisition_list')
    
    messages.warning(request, "請確認您要刪除此撥料申請單。")
    return redirect('requisition_list')


@login_required
def import_materials_to_requisition(request):
    if request.method != 'POST':
        return redirect('work_order_material_list')

    is_admin = request.user.is_superuser
    is_material_handler = request.user.groups.filter(name='撥料人員').exists()

    if not is_material_handler and not is_admin:
        messages.error(request, "您沒有權限匯入物料到撥料單。")
        return redirect('work_order_material_list')

    material_ids = request.POST.getlist('material_ids')
    requisition_id = request.POST.get('requisition_id')
    order_number = request.POST.get('order_number')

    if not material_ids or not requisition_id:
        messages.error(request, "請至少選擇一個物料和一個目標撥料單。")
        return redirect('work_order_material_list')

    try:
        requisition = Requisition.objects.get(pk=requisition_id)
        materials_to_import = WorkOrderMaterial.objects.filter(pk__in=material_ids)

        with transaction.atomic():
            new_version = MaterialListVersion.objects.create(
                requisition=requisition,
                uploaded_by=request.user
            )

            items_to_create = []
            for material in materials_to_import:
                items_to_create.append(
                    RequisitionItem(
                        material_list_version=new_version,
                        source_material=material,
                        order_number=material.order_number,
                        material_number=material.material_number,
                        item_name=material.item_name,
                        required_quantity=material.required_quantity,
                        stock_quantity=0,
                        confirmed_quantity=None,
                        is_signed_off=False,
                    )
                )
            
            RequisitionItem.objects.bulk_create(items_to_create)

            requisition.current_material_list_version = new_version
            requisition.status = 'pending'
            requisition.material_confirmed_by = None
            requisition.material_confirmed_date = None
            requisition.sign_off_by = None
            requisition.sign_off_date = None
            requisition.save()

        messages.success(request, f"成功將 {len(items_to_create)} 筆物料匯入到撥料單 '{requisition.process_type}'。")
        return redirect(f"{reverse('work_order_material_list')}?order_number={order_number}")

    except Requisition.DoesNotExist:
        messages.error(request, "找不到指定的撥料單。")
    except Exception as e:
        messages.error(request, f"匯入物料時發生錯誤: {e}")

    return redirect('work_order_material_list')


from django.db.models import Count, OuterRef, Subquery


@login_required
def work_order_material_list(request):
    is_admin = request.user.is_superuser
    is_applicant = request.user.groups.filter(name='申請人員').exists()
    is_material_handler = request.user.groups.filter(name='撥料人員').exists()

    if not is_admin and not is_applicant and not is_material_handler:
        messages.error(request, "您沒有權限查看此頁面。")
        return redirect('homepage')

    # Initialize all variables that will be used in the context
    order_number = request.GET.get('order_number', None)
    sort_by = request.GET.get('sort_by', 'material_number')
    order = request.GET.get('order', 'asc')
    process_type_filter_id = request.GET.get('process_type', None) # This is for filtering the materials displayed by ID
    process_type_name_filter = request.GET.get('process_type_name', None) # This is what comes from requisition_list
    show_inactive = request.GET.get('show_inactive', 'false').lower() == 'true' # New filter parameter

    materials = WorkOrderMaterial.objects.none()
    requisitions_for_import = Requisition.objects.none()
    process_type_choices = []
    machine_models_for_display = []
    order_numbers = WorkOrderMaterial.objects.values_list('order_number', flat=True).distinct()

    selected_process_type_for_context = None # Initialize for context

    if order_number:
        # Subquery to get storage_bin and stock_quantity from Inventory
        inventory_subquery_storage_bin = Subquery(
            Inventory.objects.filter(material_number=OuterRef('material_number')).values('storage_bin')[:1]
        )
        inventory_subquery_stock_quantity = Subquery(
            Inventory.objects.filter(material_number=OuterRef('material_number')).values('stock_quantity')[:1]
        )

        materials = WorkOrderMaterial.objects.filter(order_number=order_number).select_related('process_type').annotate(
            import_count=Count('requisition_items'),
            storage_bin=inventory_subquery_storage_bin,
            stock_quantity=inventory_subquery_stock_quantity
        )
        # Apply is_active filter
        if not show_inactive:
            materials = materials.filter(is_active=True)

        # If process_type_name_filter is provided (from requisition_list), find its ID
        if process_type_name_filter:
            try:
                # Find a WorkOrderMaterial for this order and process type name to get the machine model
                sample_material = WorkOrderMaterial.objects.filter(
                    order_number=order_number,
                    process_type__name=process_type_name_filter
                ).first()

                if sample_material and sample_material.machine_model:
                    process_type_obj = ProcessType.objects.get(
                        name=process_type_name_filter,
                        machine_model=sample_material.machine_model
                    )
                    process_type_filter_id = str(process_type_obj.id) # Use this ID for filtering and context
                else:
                    process_type_filter_id = None # No matching process type found with a machine model
            except ProcessType.DoesNotExist:
                process_type_filter_id = None # No matching process type found
            except ProcessType.MultipleObjectsReturned:
                messages.error(request, "系統錯誤：找到多個相同的投料點名稱和機型組合。")
                process_type_filter_id = None

        # Build process type choices from the unfiltered materials for this order
        all_materials_for_order = WorkOrderMaterial.objects.filter(order_number=order_number)
        unique_process_type_ids = all_materials_for_order.values_list('process_type__id', flat=True).distinct()
        unique_process_types = ProcessType.objects.filter(id__in=unique_process_type_ids).order_by('name')
        process_type_choices = []
        seen_names = set()
        for pt in unique_process_types:
            if pt.name not in seen_names:
                process_type_choices.append((pt.id, str(pt)))
                seen_names.add(pt.name)

        # Apply process type filter if provided (now process_type_filter_id holds the ID)
        if process_type_filter_id:
            materials = materials.filter(process_type__id=process_type_filter_id)
            selected_process_type_for_context = process_type_filter_id # Set for context

        # Sorting logic
        sort_mapping = {
            'material_number': 'material_number',
            'item_name': 'item_name',
            'required_quantity': 'required_quantity',
            'process_type': 'process_type__name',
            'confirmed_quantity': 'confirmed_quantity',
            'is_signed_off': 'is_signed_off',
        }
        model_sort_by = sort_mapping.get(sort_by, 'material_number')
        order_field = f'{'-' if order == "desc" else ""}{model_sort_by}'
        materials = materials.order_by(order_field)

        # Get other data needed for the context
        requisitions_for_import = Requisition.objects.filter(
            order_number=order_number,
            status__in=['pending', 'materials_confirmed', 'completed']
        ).order_by('process_type')

        unique_machine_model_ids = materials.values_list('machine_model__id', flat=True).distinct()
        unique_machine_models = MachineModel.objects.filter(id__in=unique_machine_model_ids).order_by('name')
        machine_models_for_display = [str(mm) for mm in unique_machine_models]

    query_params = request.GET.copy()
    if 'page' in query_params:
        del query_params['page']

    requisition = None
    if order_number:
        requisition = Requisition.objects.filter(order_number=order_number).first()

    context = {
        'materials': materials,
        'order_numbers': order_numbers,
        'selected_order': order_number,
        'requisitions_for_import': requisitions_for_import,
        'sort_by': sort_by,
        'order': order,
        'process_type_choices': process_type_choices,
        'selected_process_type': selected_process_type_for_context, # Use the ID for the hidden input
        'machine_models_for_display': machine_models_for_display,
        'query_params': query_params.urlencode(),
        'requisition': requisition, # Pass the requisition object
        'show_inactive': show_inactive, # New context variable
    }
    return render(request, 'requisitions/work_order_material_list.html', context)

@login_required
def shortage_materials_list(request):
    # Only allow superusers or material handlers to view this page
    if not request.user.is_superuser and not request.user.groups.filter(name='撥料人員').exists():
        messages.error(request, "您沒有權限查看此頁面。")
        return redirect('homepage')

    # Filter for active materials where required_quantity > confirmed_quantity
    # Also annotate with the calculated shortage_quantity
    shortage_materials = WorkOrderMaterial.objects.filter(
        is_active=True,
        required_quantity__gt=F('confirmed_quantity')
    ).annotate(
        shortage_quantity=ExpressionWrapper(
            F('required_quantity') - Coalesce(F('confirmed_quantity'), 0),
            output_field=DecimalField()
        )
    ).order_by('order_number', 'material_number').distinct()

    context = {
        'shortage_materials': shortage_materials,
    }
    return render(request, 'requisitions/shortage_materials_list.html', context)



@login_required
def update_process_type_db(request):
    if not request.user.is_superuser:
        messages.error(request, "您沒有權限執行此操作。")
        return redirect('homepage')

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
                return redirect('homepage')
            except Exception as e:
                messages.error(request, f"更新資料庫時發生錯誤: {e}")
    else:
        form = UpdateProcessTypeDBForm()
    
    return render(request, 'requisitions/update_process_type_db.html', {'form': form})


@login_required
def upload_inventory_data(request): # This will now be for stock quantity only
    if not request.user.is_superuser:
        messages.error(request, "您沒有權限執行此操作。")
        return redirect('homepage')

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
                return redirect('homepage')

            except Exception as e:
                messages.error(request, f"上傳檔案時發生錯誤: {e}")
                import traceback
                print(traceback.format_exc())
    else:
        form = UploadInventoryFileForm()
    
    return render(request, 'requisitions/upload_inventory_data.html', {'form': form})


@login_required
def upload_storage_bin_data(request):
    if not request.user.is_superuser:
        messages.error(request, "您沒有權限執行此操作。")
        return redirect('homepage')

    if request.method == 'POST':
        form = UploadStorageBinFileForm(request.POST, request.FILES)
        if form.is_valid():
            excel_file = request.FILES['file']
            try:
                df = pd.read_excel(excel_file, dtype=str)
                df.columns = df.columns.str.strip()

                if '物料' not in df.columns:
                    raise ValueError("Excel 檔案中找不到 '物料' 欄位。")
                if '儲格' not in df.columns:
                    raise ValueError("Excel 檔案中找不到 '儲格' 欄位。")

                updated_count = 0
                created_count = 0

                with transaction.atomic():
                    for index, row in df.iterrows():
                        material_number = row.get('物料')
                        storage_bin = row.get('儲格', '')

                        if not material_number:
                            messages.warning(request, f"跳過第 {index+2} 行: 物料為空。")
                            continue

                        inventory_item, created = Inventory.objects.get_or_create(
                            material_number=material_number,
                            defaults={
                                'storage_bin': storage_bin, # Set storage_bin for new objects
                                'stock_quantity': 0, # Default stock_quantity for new objects
                            }
                        )
                        if not created:
                            # If object already existed, only update storage_bin
                            inventory_item.storage_bin = storage_bin
                            inventory_item.save()
                        if created:
                            created_count += 1
                        else:
                            updated_count += 1
                
                messages.success(request, f"儲格資料上傳成功！新增 {created_count} 筆，更新 {updated_count} 筆。")
                return redirect('homepage')

            except Exception as e:
                messages.error(request, f"上傳檔案時發生錯誤: {e}")
                import traceback
                print(traceback.format_exc())
    else:
        form = UploadStorageBinFileForm()
    
    return render(request, 'requisitions/upload_storage_bin_data.html', {'form': form})


@login_required
def view_inventory_database(request):
    if not request.user.is_superuser:
        messages.error(request, "您沒有權限執行此操作。")
        return redirect('homepage')
    inventory_items = Inventory.objects.all()
    context = {
        'inventory_items': inventory_items,
    }
    return render(request, 'requisitions/view_inventory_database.html', context)


@login_required
def sign_off_item(request, pk, version_pk, item_pk):
    requisition = get_object_or_404(Requisition, pk=pk)
    target_version = get_object_or_404(MaterialListVersion, pk=version_pk, requisition=requisition)
    item = get_object_or_404(RequisitionItem, pk=item_pk, material_list_version=target_version)

    is_admin = request.user.is_superuser
    is_applicant = request.user.groups.filter(name='申請人員').exists()

    if not is_applicant and not is_admin:
        return JsonResponse({'success': False, 'message': "您沒有權限簽收物料項目。"}, status=403)

    if request.method == 'POST':
        if not item.is_signed_off:
            item.is_signed_off = True
            item.save()

            if item.source_material:
                item.source_material.is_signed_off = True
                if item.confirmed_quantity is not None:
                    item.source_material.confirmed_quantity = item.confirmed_quantity
                item.source_material.save()

            messages.success(request, f"物料項目 '{item.item_name}' 已成功簽收。")

            all_items_in_version = RequisitionItem.objects.filter(material_list_version=target_version)
            if all(i.is_signed_off for i in all_items_in_version):
                if target_version == requisition.current_material_list_version:
                    with transaction.atomic():
                        requisition.status = 'completed'
                        requisition.sign_off_by = request.user
                        requisition.sign_off_date = timezone.now()
                        requisition.save()
                        messages.success(request, "撥料申請單已全部最終簽收！")
                else:
                    messages.info(request, f"物料清單版本 '{target_version.uploaded_at.strftime('%Y-%m-%d %H:%M')}' 已全部最終簽收。")
            return JsonResponse({'success': True, 'message': "物料項目已成功簽收。"})
        else:
            messages.info(request, "此物料項目已簽收。")
            return JsonResponse({'success': False, 'message': "此物料項目已簽收。"})

    return JsonResponse({'success': False, 'message': "無效的請求方法。"}, status=405)

@login_required
def update_material_process_type(request, material_id):
    if not request.user.is_superuser:
        return JsonResponse({'success': False, 'message': '權限不足'}, status=403)

    if request.method == 'POST':
        try:
            material = get_object_or_404(WorkOrderMaterial, pk=material_id)
            data = json.loads(request.body)
            new_process_type_id = data.get('process_type')

            if not new_process_type_id:
                return JsonResponse({'success': False, 'message': '未提供投料點 ID'}, status=400)

            # Get the ProcessType instance from the provided ID
            process_type_instance = get_object_or_404(ProcessType, pk=new_process_type_id)

            # Assign the actual model instance to the ForeignKey field
            material.process_type = process_type_instance
            material.save()
            
            return JsonResponse({'success': True, 'message': '投料點更新成功'})

        except WorkOrderMaterial.DoesNotExist:
            return JsonResponse({'success': False, 'message': '找不到物料'}, status=404)
        except ProcessType.DoesNotExist:
            return JsonResponse({'success': False, 'message': '找不到指定的投料點'}, status=404)
        except Exception as e:
            import traceback
            return JsonResponse({'success': False, 'message': traceback.format_exc()}, status=500)

    return JsonResponse({'success': False, 'message': '無效的請求'}, status=400)


@login_required
def get_process_types_for_model(request):
    machine_model_id = request.GET.get('machine_model_id')
    
    if not machine_model_id or not machine_model_id.isdigit():
        return JsonResponse([], safe=False)
    
    try:
        process_types = ProcessType.objects.filter(machine_model_id=int(machine_model_id)).values('id', 'name')
        return JsonResponse(list(process_types), safe=False)
    except Exception as e:
        import traceback
        return JsonResponse({'error': traceback.format_exc()}, status=500)


@login_required
def process_types_management(request):
    if not request.user.is_superuser:
        messages.error(request, "您沒有權限執行此操作。")
        return redirect('homepage')

    process_types = ProcessType.objects.all().select_related('machine_model')
    form = ProcessTypeForm()

    if request.method == 'POST':
        if 'add_process_type' in request.POST:
            form = ProcessTypeForm(request.POST)
            if form.is_valid():
                try:
                    form.save()
                    messages.success(request, "投料點新增成功！")
                    return redirect('process_types_management')
                except IntegrityError:
                    messages.error(request, "該機型下已存在同名的投料點，請檢查。")
                except Exception as e:
                    messages.error(request, f"新增投料點時發生錯誤: {e}")
            else:
                messages.error(request, "表單驗證失敗，請檢查輸入。")
        elif 'edit_process_type' in request.POST:
            process_type_id = request.POST.get('process_type_id')
            process_type_instance = get_object_or_404(ProcessType, pk=process_type_id)
            form = ProcessTypeForm(request.POST, instance=process_type_instance)
            if form.is_valid():
                try:
                    form.save()
                    messages.success(request, "投料點更新成功！")
                    return redirect('process_types_management')
                except IntegrityError:
                    messages.error(request, "該機型下已存在同名的投料點，請檢查。")
                except Exception as e:
                    messages.error(request, f"更新投料點時發生錯誤: {e}")
            else:
                messages.error(request, "表單驗證失敗，請檢查輸入。")
        elif 'delete_process_type' in request.POST:
            process_type_id = request.POST.get('process_type_id')
            process_type_instance = get_object_or_404(ProcessType, pk=process_type_id)
            try:
                process_type_instance.delete()
                messages.success(request, "投料點刪除成功！")
                return redirect('process_types_management')
            except Exception as e:
                messages.error(request, f"刪除投料點時發生錯誤: {e}")

    context = {
        'process_types': process_types,
        'form': form,
        'machine_models': MachineModel.objects.all(), # Pass all machine models for the form
    }
    return render(request, 'requisitions/process_types_management.html', context)


@login_required
def get_requisition_details_json(request, pk):
    requisition = get_object_or_404(Requisition, pk=pk)

    # Get user roles for conditional rendering in frontend
    is_admin = request.user.is_superuser
    is_applicant = request.user.groups.filter(name='申請人員').exists()
    is_material_handler = request.user.groups.filter(name='撥料人員').exists()

    data = {
        'pk': requisition.pk,
        'order_number': requisition.order_number,
        'applicant_name': requisition.applicant.get_full_name(),
        'applicant_id': requisition.applicant.id,
        'request_date': requisition.request_date.strftime('%Y-%m-%d'),
        'process_type': requisition.process_type,
        'status': requisition.status,
        'status_display': requisition.get_status_display(), # Use get_status_display for verbose status
        'created_at': requisition.created_at.strftime('%Y-%m-%d %H:%M'),
        'remarks': requisition.remarks,
        'is_admin': is_admin,
        'is_applicant': is_applicant,
        'is_material_handler': is_material_handler,
    }
    return JsonResponse(data)


@login_required
def get_requisition_images_json(request, pk):
    requisition = get_object_or_404(Requisition, pk=pk)
    images = requisition.images.all().order_by('-uploaded_at') # Assuming 'images' related_name on RequisitionImage model

    images_data = []
    for image in images:
        images_data.append({
            'url': image.image.url,
            'uploaded_at': image.uploaded_at.strftime('%Y-%m-%d %H:%M'),
            'uploaded_by': image.uploaded_by.username if image.uploaded_by else 'N/A',
        })
    return JsonResponse({'images': images_data})


@login_required
def get_requisition_items_json(request, pk):
    requisition = get_object_or_404(Requisition, pk=pk)
    
    if not requisition.current_material_list_version:
        return JsonResponse({'items': []})

    items = RequisitionItem.objects.filter(
        material_list_version=requisition.current_material_list_version
    ).select_related('source_material__machine_model')

    items_data = []
    for item in items:
        items_data.append({
            'pk': item.pk,
            'material_number': item.material_number,
            'item_name': item.item_name,
            'machine_model': item.source_material.machine_model.name if item.source_material and item.source_material.machine_model else 'N/A',
            'required_quantity': item.required_quantity,
            'stock_quantity': item.stock_quantity,
            'storage_bin': item.storage_bin,
            'confirmed_quantity': item.confirmed_quantity,
            'is_signed_off': item.is_signed_off,
        })
    
    return JsonResponse({'items': items_data})


@login_required
def supplement_material(request, pk):
    requisition = get_object_or_404(Requisition, pk=pk)

    # Ensure only authorized users can supplement
    is_admin = request.user.is_superuser
    is_material_handler = request.user.groups.filter(name='撥料人員').exists()
    if not is_admin and not is_material_handler:
        return redirect('requisition_detail', pk=pk)

    # Prevent supplementing a completed requisition
    if requisition.status == 'completed':
        messages.error(request, "無法為已完成的申請單補料。")
        return redirect('requisition_detail', pk=pk)

    # Get materials already in the current version
    current_material_pks = []
    if requisition.current_material_list_version:
        current_material_pks = RequisitionItem.objects.filter(
            material_list_version=requisition.current_material_list_version
        ).values_list('source_material__pk', flat=True)

    # Get all possible materials for this order, excluding those already in the list
    available_materials = WorkOrderMaterial.objects.filter(
        order_number=requisition.order_number
    ).exclude(
        pk__in=list(current_material_pks)
    )

    if request.method == 'POST':
        selected_material_ids = request.POST.getlist('material_ids')
        if not selected_material_ids:
            messages.error(request, "您沒有選擇任何物料。")
            return redirect('supplement_material', pk=pk)

        try:
            with transaction.atomic():
                # Create a new version for the supplemented list
                new_version = MaterialListVersion.objects.create(
                    requisition=requisition,
                    uploaded_by=request.user
                )

                # 1. Copy items from the old version (if it exists)
                if requisition.current_material_list_version:
                    old_items = RequisitionItem.objects.filter(material_list_version=requisition.current_material_list_version)
                    for item in old_items:
                        RequisitionItem.objects.create(
                            material_list_version=new_version,
                            source_material=item.source_material,
                            order_number=item.order_number,
                            material_number=item.material_number,
                            item_name=item.item_name,
                            required_quantity=item.required_quantity,
                            stock_quantity=item.stock_quantity,
                            confirmed_quantity=item.confirmed_quantity, # Preserve confirmed quantity from old version
                            is_signed_off=item.is_signed_off # Preserve sign-off status
                        )

                # 2. Add the new supplemented materials
                materials_to_add = WorkOrderMaterial.objects.filter(pk__in=selected_material_ids)
                for material in materials_to_add:
                    quantity = request.POST.get(f'quantity_{material.pk}')
                    if quantity:
                        RequisitionItem.objects.create(
                            material_list_version=new_version,
                            source_material=material,
                            order_number=material.order_number,
                            material_number=material.material_number,
                            item_name=material.item_name,
                            required_quantity=quantity,
                            stock_quantity=0,  # New items from supplement have 0 stock qty by default
                            confirmed_quantity=None,
                            is_signed_off=False
                        )

                # 3. Update the requisition to point to the new version and reset status
                requisition.current_material_list_version = new_version
                requisition.status = 'pending'
                requisition.material_confirmed_by = None
                requisition.material_confirmed_date = None
                requisition.sign_off_by = None
                requisition.sign_off_date = None
                requisition.save()

                messages.success(request, f"成功補料 {len(selected_material_ids)} 項，申請單狀態已重置為待撥料。")
                return redirect('requisition_detail', pk=pk)

        except Exception as e:
            messages.error(request, f"補料時發生錯誤: {e}")
            return redirect('supplement_material', pk=pk)

    context = {
        'requisition': requisition,
        'available_materials': available_materials,
    }
    return render(request, 'requisitions/supplement_material.html', context)


@login_required
def upload_requisition_images(request, pk):
    requisition = get_object_or_404(Requisition, pk=pk)

    if request.method == 'POST':
        image_form = RequisitionImageUploadForm(request.POST, request.FILES)
        if image_form.is_valid():
            uploaded_count = 0
            for image_file in request.FILES.getlist('images'):
                RequisitionImage.objects.create(
                    requisition=requisition,
                    image=image_file,
                    uploaded_by=request.user
                )
                uploaded_count += 1
            if uploaded_count > 0:
                messages.success(request, f"成功上傳 {uploaded_count} 張圖片！")
            else:
                messages.info(request, "沒有圖片被上傳。")
            return redirect('view_requisition_images', pk=pk) # Redirect to prevent re-submission
        else:
            messages.error(request, "圖片上傳失敗，請檢查檔案格式。")
    return redirect('view_requisition_images', pk=pk) # Redirect if not POST or form not valid









@login_required
def upload_work_order_material_images(request, pk):
    requisition = get_object_or_404(Requisition, pk=pk)
    
    if request.method == 'POST':
        form = WorkOrderMaterialImageUploadForm(request.POST, request.FILES)
        if form.is_valid():
            process_type = form.cleaned_data.get('process_type')
            for image_file in request.FILES.getlist('images'):
                WorkOrderMaterialImage.objects.create(
                    requisition=requisition,
                    process_type=process_type,
                    image=image_file,
                    uploaded_by=request.user
                )
            messages.success(request, "圖片上傳成功！")
    return redirect(request.META.get('HTTP_REFERER', 'work_order_material_list'))




@login_required
@transaction.atomic
def update_work_order_quantities(request):
      if request.method != 'POST':
          return redirect('work_order_material_list')

      order_number = request.POST.get('order_number')
      process_type_filter = request.POST.get('process_type_filter', '')
      updated_materials = []

      print(f"Received POST data: {request.POST}") # DEBUG

      redirect_url = request.META.get('HTTP_REFERER', reverse('work_order_material_list'))
      query_string = ''
      if '?' in redirect_url:
          query_string = '?' + redirect_url.split('?', 1)[1]
          redirect_url = redirect_url.split('?', 1)[0]

      for key, value in request.POST.items():
          print(f"Processing key: {key}, value: {value}") # DEBUG
          if key.startswith('change_') and value:
              try:
                  material_id = int(key.split('_')[1])
                  quantity_change = Decimal(value)

                  print(f"  Parsed: material_id={material_id},quantity_change={quantity_change}") # DEBUG

                  if quantity_change == 0:
                      print("  Skipping: quantity_change is 0") # DEBUG
                      continue

                  material =WorkOrderMaterial.objects.get(pk=material_id)

                  current_confirmed = material.confirmed_quantity if material.confirmed_quantity is not None else Decimal('0')

                  new_confirmed_quantity = current_confirmed + quantity_change
                  if new_confirmed_quantity > material.required_quantity:
                      messages.warning(request, f"物料 {material.material_number} 的撥料數量 ({new_confirmed_quantity}) 超過需求數量 ({material.required_quantity})。")
                  material.confirmed_quantity = new_confirmed_quantity
                  material.save()

                  transaction_type = 'ALLOCATION' if quantity_change > 0 else 'RETURN'

                  WorkOrderMaterialTransaction.objects.create(
                      work_order_material=material,
                      user=request.user,
                      transaction_type=transaction_type,
                      quantity_change=quantity_change,
                      new_confirmed_quantity=new_confirmed_quantity,
                      notes="手動撥料/退料操作"
                  )
                  updated_materials.append(f"{material.material_number} ({quantity_change:+.2f})")

              except (ValueError, WorkOrderMaterial.DoesNotExist) as e:
                  print(f"  Error in loop: {e}") # DEBUG
                  messages.error(request, f"處理物料 ID  {key.split('_')[1]} 時發生錯誤: {e}，部分或所有變更可能未儲存。")
                  return redirect(f'{redirect_url}{query_string}')

      if updated_materials:
          if process_type_filter:
              try:
                  process_type = get_object_or_404(ProcessType, id=process_type_filter)
                  requisition = get_object_or_404(Requisition, order_number=order_number, process_type=process_type.name)
                  requisition.dispatch_performed = True
                  requisition.save()
              except Exception as e:
                  messages.error(request, f"更新撥料狀態時發生錯誤: {e}")

      print(f"Updated materials list: {updated_materials}") # DEBUG
      if updated_materials:messages.success(request, f"成功更新撥料數量: {','.join(updated_materials)}")
      else:messages.info ( request, "沒有偵測到任何數量變動。" )

      return redirect(f'{redirect_url}{query_string}')

@login_required
def generate_dispatch_note(request, pk):
    requisition = get_object_or_404(Requisition, pk=pk)
    materials = WorkOrderMaterial.objects.filter(
        order_number=requisition.order_number,
        process_type__name=requisition.process_type,
        is_active=True # Only active materials
    ).filter(confirmed_quantity__gt=0)

    # Fetch images related to this requisition and its process type
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
def generate_backorder_note(request, pk):
    requisition = get_object_or_404(Requisition, pk=pk)

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

    context = {
        'requisition': requisition,
        'shortage_materials': shortage_materials,
    }
    return render(request, 'requisitions/backorder_note.html', context)


@login_required
def export_backorder_note_excel(request, pk):
    requisition = get_object_or_404(Requisition, pk=pk)

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
def update_dispatch_note(request, pk):
    requisition = get_object_or_404(Requisition, pk=pk)
    if request.method == 'POST':
        confirm_value = request.POST.get('confirm')
        if confirm_value:
            material_id, action = confirm_value.split('_')
            try:
                material = WorkOrderMaterial.objects.get(pk=material_id)
                if action == 'yes':
                    material.confirmed_quantity = material.required_quantity
                    material.is_signed_off = True
                    messages.success(request, f"物料 {material.material_number} 已確認撥料。")
                else:
                    material.confirmed_quantity = 0
                    material.is_signed_off = False
                    messages.info(request, f"物料 {material.material_number} 已標記為未撥料。")
                material.save()
            except WorkOrderMaterial.DoesNotExist:
                messages.error(request, f"物料 ID {material_id} 不存在。")
            except Exception as e:
                messages.error(request, f"更新物料時發生錯誤: {e}")
    return redirect('generate_dispatch_note', pk=requisition.pk)


@login_required
@transaction.atomic
def update_material_dispatch_status(request, pk):
    if request.method == 'POST':
        try:
            data = json.loads(request.body)
            material_id = data.get('material_id')
            action = data.get('action') # 'yes' or 'no'

            material = get_object_or_404(WorkOrderMaterial, pk=material_id)
            requisition = get_object_or_404(Requisition, pk=pk)

            if action == 'yes':
                material.confirmed_quantity = material.required_quantity
                material.is_signed_off = True
                message = f"物料 {material.material_number} 已確認撥料。"
            elif action == 'no':
                material.confirmed_quantity = Decimal('0.00') # Set to 0 for backorder
                material.is_signed_off = False
                message = f"物料 {material.material_number} 已取消撥料並移至欠料。"
            else:
                return JsonResponse({'success': False, 'message': '無效的操作。'}, status=400)

            material.save()

            # Update Requisition status if all materials are dispatched/undispatched
            # This logic might need refinement based on exact business rules
            # For now, let's assume if all materials in the current dispatch note are handled,
            # we can update the requisition status.
            
            # For now, let's just return success and let the user refresh or handle UI updates
            return JsonResponse({'success': True, 'message': message, 'new_confirmed_quantity': str(material.confirmed_quantity), 'new_is_signed_off': material.is_signed_off})

        except WorkOrderMaterial.DoesNotExist:
            return JsonResponse({'success': False, 'message': '找不到物料。'}, status=404)
        except Requisition.DoesNotExist:
            return JsonResponse({'success': False, 'message': '找不到撥料單。'}, status=404)
        except json.JSONDecodeError:
            return JsonResponse({'success': False, 'message': '無效的 JSON 請求。'}, status=400)
        except Exception as e:
            import traceback
            return JsonResponse({'success': False, 'message': f'處理請求時發生錯誤: {e}\n{traceback.format_exc()}'}, status=500)
    return JsonResponse({'success': False, 'message': '無效的請求方法。'}, status=405)
