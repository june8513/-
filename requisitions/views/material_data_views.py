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

@login_required
def view_process_type_database(request):
    if not request.user.is_superuser:
        messages.error(request, "您沒有權限執行此操作。")
        return redirect('homepage')
    db_path = os.path.join(settings.BASE_DIR, 'output.xlsx')
    data = []
    headers = []
    try:
        excel_sheets = pd.read_excel(db_path, engine='openpyxl',
sheet_name=None)
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
    
    materials = WorkOrderMaterial.objects.all().select_related('process_type', 'machine_model').order_by('order_number', 'material_number')

    paginator = Paginator(materials, 20) # Show 20 materials per page
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)

    context = {
        'materials': page_obj,
    }
    return render(request, 'requisitions/view_database.html', context)


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

from collections import defaultdict
from django.db.models import Sum

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
        # Subquery to get bin from Material model
        material_subquery_bin = Subquery(
            Material.objects.filter(material_code=OuterRef('material_number')).values('bin')[:1]
        )
        # Subquery to get stock_quantity from Inventory
        inventory_subquery_stock_quantity = Subquery(
            Inventory.objects.filter(material_number=OuterRef('material_number')).values('stock_quantity')[:1]
        )

        materials = WorkOrderMaterial.objects.filter(order_number=order_number).select_related('process_type').annotate(
            import_count=Count('requisition_items'),
            bin=material_subquery_bin, # Fetch bin from Material model
            stock_quantity=inventory_subquery_stock_quantity
        )
        # Apply is_active filter
        if not show_inactive:
            materials = materials.filter(is_active=True)
        
        # DEBUG: Print bin values
        for m in materials:
            print(f"Material: {m.material_number}, Bin: {m.bin}")

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
            'bin': 'bin', # Add bin for sorting
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
def archived_work_order_material_list(request):
    is_admin = request.user.is_superuser
    is_applicant = request.user.groups.filter(name='申請人員').exists()
    is_material_handler = request.user.groups.filter(name='撥料人員').exists()

    if not is_admin and not is_applicant and not is_material_handler:
        messages.error(request, "您沒有權限查看此頁面。")
        return redirect('homepage')

    order_number = request.GET.get('order_number', None)
    sort_by = request.GET.get('sort_by', 'material_number')
    order = request.GET.get('order', 'asc')
    process_type_filter_id = request.GET.get('process_type', None)

    materials = WorkOrderMaterial.objects.none()
    process_type_choices = []
    machine_models_for_display = []
    order_numbers = WorkOrderMaterial.objects.values_list('order_number', flat=True).distinct()

    selected_process_type_for_context = None

    if order_number:
        inventory_subquery_storage_bin = Subquery(
            Inventory.objects.filter(material_number=OuterRef('material_number')).values('storage_bin')[:1]
        )
        inventory_subquery_stock_quantity = Subquery(
            Inventory.objects.filter(material_number=OuterRef('material_number')).values('stock_quantity')[:1]
        )

        materials = WorkOrderMaterial.objects.filter(
            order_number=order_number,
            is_active=False # Filter for inactive materials
        ).select_related('process_type').annotate(
            import_count=Count('requisition_items'),
            storage_bin=inventory_subquery_storage_bin,
            stock_quantity=inventory_subquery_stock_quantity
        )

        all_materials_for_order = WorkOrderMaterial.objects.filter(order_number=order_number, is_active=False)
        unique_process_type_ids = all_materials_for_order.values_list('process_type__id', flat=True).distinct()
        unique_process_types = ProcessType.objects.filter(id__in=unique_process_type_ids).order_by('name')
        process_type_choices = []
        seen_names = set()
        for pt in unique_process_types:
            if pt.name not in seen_names:
                process_type_choices.append((pt.id, str(pt)))
                seen_names.add(pt.name)

        if process_type_filter_id:
            materials = materials.filter(process_type__id=process_type_filter_id)
            selected_process_type_for_context = process_type_filter_id

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

        unique_machine_model_ids = materials.values_list('machine_model__id', flat=True).distinct()
        unique_machine_models = MachineModel.objects.filter(id__in=unique_machine_model_ids).order_by('name')
        machine_models_for_display = [str(mm) for mm in unique_machine_models]

    query_params = request.GET.copy()
    if 'page' in query_params:
        del query_params['page']

    context = {
        'materials': materials,
        'order_numbers': order_numbers,
        'selected_order': order_number,
        'sort_by': sort_by,
        'order': order,
        'process_type_choices': process_type_choices,
        'selected_process_type': selected_process_type_for_context,
        'machine_models_for_display': machine_models_for_display,
        'query_params': query_params.urlencode(),
    }
    return render(request, 'requisitions/archived_work_order_material_list.html', context)

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
      return redirect(f'{redirect_url}{query_string}')

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