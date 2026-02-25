from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.utils import timezone
from ..forms import RequisitionForm, UploadFileForm, OrderModelUploadForm, MaterialDetailsUploadForm, RequisitionItemMaterialConfirmationFormSet, RequisitionItemSignOffFormSet, UpdateProcessTypeDBForm, UploadInventoryFileForm, ProcessTypeForm, RequisitionImageForm, WorkOrderMaterialImageUploadForm
from ..models import Requisition, RequisitionItem, WorkOrderMaterial, Inventory, MachineModel, ProcessType, RequisitionImage, WorkOrderMaterialTransaction, WorkOrderMaterialImage, WorkOrder
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
from requisitions.utils import process_order_model_excel, process_material_details_excel, notify_requisition_shortages
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

    if not (is_admin or is_material_handler or is_applicant): # Allow applicant
        messages.error(request, "您沒有權限查看此頁面。")
        return redirect('core:homepage')

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
    order_numbers = WorkOrder.objects.filter(is_archived=False).values_list('order_number', flat=True).order_by('-order_number')
    
    all_process_type_names = ['機械', '系統', '電裝', '鑄件', '護蓋', '刀庫', '出貨', '組件']

    # Check if the requested order_number is archived
    if order_number:
        try:
            work_order = WorkOrder.objects.get(order_number=order_number)
            if work_order.is_archived:
                messages.error(request, f"工單 {order_number} 已被歸檔，無法在此頁面查看。")
                return redirect('requisitions:work_order_list')
        except WorkOrder.DoesNotExist:
            messages.error(request, f"找不到工單 {order_number}。")
            return redirect('requisitions:work_order_list')

    selected_process_type_for_context = None # Initialize for context

    if order_number:
        # Subquery to get stock_quantity from the main inventory.Material model
        inventory_subquery_stock_quantity = Subquery(
            Material.objects.filter(material_code=OuterRef('material_number')).values('system_quantity')[:1]
        )
        material_subquery_bin = Subquery(
            Material.objects.filter(material_code=OuterRef('material_number')).values('bin')[:1]
        )

        # Subquery to calculate the total confirmed quantity from all related RequisitionItems
        total_confirmed_subquery = RequisitionItem.objects.filter(
            source_material=OuterRef('pk')
        ).values('source_material').annotate(
            total=Sum('confirmed_quantity')
        ).values('total')

        materials = WorkOrderMaterial.objects.filter(order_number=order_number).select_related('process_type').annotate(
            import_count=Count('requisition_items'),
            bin=material_subquery_bin,
            stock_quantity=inventory_subquery_stock_quantity,
            total_confirmed_quantity=Coalesce(Subquery(total_confirmed_subquery, output_field=DecimalField()), Decimal('0.0'))
        )
        # Apply is_active filter
        if not show_inactive:
            # Show active materials OR inactive materials that have confirmed quantity > 0
            materials = materials.filter(Q(is_active=True) | Q(confirmed_quantity__gt=0))
        


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
            'confirmed_quantity': 'total_confirmed_quantity', # Use the annotated field for sorting
            'is_signed_off': 'is_signed_off',
            'bin': 'bin', # Add bin for sorting
        }
        model_sort_by = sort_mapping.get(sort_by, 'material_number')
        order_field = f'{'-' if order == "desc" else ""}{model_sort_by}'
        materials = materials.order_by(order_field)

        # --- Check for Earlier Shortages (Queue Jumping Alert) ---
        backlog_map = {}
        current_req_dates = {}
        
        if order_number:
            # Get dates for current order's requisitions
            current_reqs = Requisition.objects.filter(order_number=order_number)
            for req in current_reqs:
                current_req_dates[req.process_type] = req.request_date

        if order_number and materials:
            target_material_numbers = list(materials.values_list('material_number', flat=True))
            
            if current_req_dates:
                 # Find all *other* active shortages for these materials
                 other_shortages = WorkOrderMaterial.objects.filter(
                    material_number__in=target_material_numbers,
                    is_active=True,
                    required_quantity__gt=Coalesce(F('confirmed_quantity'), 0)
                ).exclude(
                    order_number=order_number
                ).select_related('process_type')
                
                 shortage_groups = {}
                 for s in other_shortages:
                     p_name = s.process_type.name if s.process_type else None
                     key = (s.order_number, p_name)
                     if key not in shortage_groups: shortage_groups[key] = []
                     shortage_groups[key].append(s)
                     
                 if shortage_groups:
                    date_q = Q()
                    for (o_num, p_name) in shortage_groups.keys():
                        if p_name: date_q |= Q(order_number=o_num, process_type=p_name)
                        else: date_q |= Q(order_number=o_num, process_type__isnull=True)
                    
                    other_reqs = Requisition.objects.filter(
                        date_q,
                        status__in=['demand_submitted', 'dispatch_in_progress'],
                        is_archived=False
                    ).values('order_number', 'process_type', 'request_date')
                    
                    other_req_map = { (r['order_number'], r['process_type']): r['request_date'] for r in other_reqs }
                    
                    for s_list in shortage_groups.values():
                        for s in s_list:
                            p_name = s.process_type.name if s.process_type else None
                            other_date = other_req_map.get((s.order_number, p_name))
                            if other_date:
                                if s.material_number not in backlog_map: backlog_map[s.material_number] = []
                                backlog_map[s.material_number].append({
                                    'order': s.order_number,
                                    'date': other_date,
                                    'shortage': s.required_quantity - (s.confirmed_quantity or 0)
                                })
        
        # Convert queryset to list and attach info
        materials_list = list(materials)
        for m in materials_list:
            m_process_name = m.process_type.name if m.process_type else None
            current_date = current_req_dates.get(m_process_name)
            
            m.backlog_info = []
            if current_date and m.material_number in backlog_map:
                m.backlog_info = [b for b in backlog_map[m.material_number] if b['date'] < current_date]
        
        # Use the list for context
        materials = materials_list 

        # Get other data needed for the context
        unique_machine_model_ids = set(m.machine_model_id for m in materials if m.machine_model_id)
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
        'is_admin': is_admin, # Pass to context
        'is_material_handler': is_material_handler, # Pass to context
        'is_applicant': is_applicant, # Pass to context
        'all_process_type_names': all_process_type_names,
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
    order_numbers = WorkOrder.objects.filter(is_archived=True).values_list('order_number', flat=True).order_by('-order_number')

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
          return redirect('requisitions:work_order_material_list')

      order_number = request.POST.get('order_number')
      process_type_filter = request.POST.get('process_type_filter', '')
      updated_materials = [] # Re-initialize updated_materials here
      redirect_to_requisition_pk = None # Initialize a variable to store the PK for redirection
      affected_requisition_ids = set() # Track affected requisitions for notification

      print(f"Received POST data: {request.POST}") # DEBUG

      redirect_url = request.META.get('HTTP_REFERER', reverse('requisitions:work_order_material_list'))
      query_string = ''
      if '?' in redirect_url:
          query_string = '?' + redirect_url.split('?', 1)[1]
          redirect_url = redirect_url.split('?', 1)[0]

      # Check if process_type_filter is valid
      if not process_type_filter:
          messages.error(request, "請先選擇投料點再進行操作。")
          return redirect(f'{redirect_url}{query_string}')

      # Find the ProcessType object by its ID
      process_type_obj = get_object_or_404(ProcessType, pk=process_type_filter)

      # Get the unique Requisition associated with this order and process type
      current_requisition = Requisition.objects.filter(
          order_number=order_number,
          process_type=process_type_obj.name # ProcessType name as CharField
      ).first()
      
      if not current_requisition:
          messages.error(request, f"找不到訂單 {order_number} 和流程 {process_type_obj.name} 對應的撥料申請單。")
          return redirect(f'{redirect_url}{query_string}')

      # Get all WorkOrderMaterials relevant to this order and process type
      all_relevant_work_order_materials = WorkOrderMaterial.objects.filter(
          order_number=order_number,
          process_type=process_type_obj
      )
      processed_work_order_material_pks = set() # To track materials explicitly handled by user input

      for key, value in request.POST.items():
          print(f"Processing key: {key}, value: {value}") # DEBUG
          
          if key.startswith('processtype_'):
              try:
                  material_id = int(key.split('_')[1])
                  new_pt_name = value
                  
                  if not new_pt_name: continue
                  
                  material = WorkOrderMaterial.objects.get(pk=material_id)
                  current_pt_name = material.process_type.name if material.process_type else None
                  
                  if current_pt_name != new_pt_name:
                      # 使用 get_or_create 來自動建立不存在的投料點
                      new_pt_obj, created = ProcessType.objects.get_or_create(
                          name=new_pt_name, 
                          machine_model=material.machine_model
                      )
                      if created:
                          messages.info(request, f"已為機型 {material.machine_model} 自動建立新投料點 '{new_pt_name}'。")
                      
                      material.process_type = new_pt_obj
                      material.save()
                      updated_materials.append(f"物料 {material.material_number} 投料點更新: {new_pt_name}")
                      
                      # Learning
                      from requisitions.models import MaterialProcessTypeRule
                      material_prefix = material.material_number[:10]
                      MaterialProcessTypeRule.objects.update_or_create(
                          material_prefix=material_prefix,
                          machine_model_name=material.machine_model.name,
                          defaults={
                              'process_type_name': new_pt_name,
                              'updated_by': request.user
                          }
                      )
              except Exception as e:
                  print(f"Error updating PT: {e}")

          if key.startswith('change_') and value:
              try:
                  material_id = int(key.split('_')[1])
                  quantity_change = Decimal(value)

                  print(f"  Parsed: material_id={material_id},quantity_change={quantity_change}") # DEBUG

                  if quantity_change == 0:
                      print("  Skipping: quantity_change is 0") # DEBUG
                      # Even if quantity_change is 0, we treat it as "processed" by user action
                      # So, if user explicitly set to 0, it means they considered it.
                      # This ensures it's not picked up by the 'unprocessed' loop later.
                      material = WorkOrderMaterial.objects.get(pk=material_id)
                      processed_work_order_material_pks.add(material.pk) 
                      continue

                  material = WorkOrderMaterial.objects.get(pk=material_id)
                  processed_work_order_material_pks.add(material.pk) # Mark as processed


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

                  # --- New Logic for Requisition and RequisitionItem ---
                  # Find relevant Requisitions
                  relevant_requisitions = Requisition.objects.filter(
                      order_number=material.order_number,
                      process_type=material.process_type.name,
                      status__in=['demand_submitted', 'dispatch_in_progress', 'dispatch_completed', 'signed_off'] # Include more statuses
                  )

                  for req in relevant_requisitions:
                      print(f"DEBUG: Processing Requisition (PK: {req.pk}, Order: {req.order_number}, Process: {req.process_type}, Status: {req.status}) for WorkOrderMaterial (PK: {material.pk}, Material: {material.material_number})") # DEBUG

                      # Update Requisition status if it's still 'demand_submitted'
                      if req.status == 'demand_submitted':
                          req.status = 'dispatch_in_progress'
                          req.save()

                      # Create or update RequisitionItem
                      requisition_item, created = RequisitionItem.objects.update_or_create(
                          requisition=req,
                          material_number=material.material_number, # Use material_number for lookup
                          defaults={
                              'source_material': material, # Keep source_material in defaults
                              'order_number': material.order_number,
                              'item_name': material.item_name,
                              'required_quantity': material.required_quantity,
                              'stock_quantity': Material.objects.filter(material_code=material.material_number).values_list('system_quantity', flat=True).first() or Decimal('0'), # Fetch current stock from the correct source
                              'confirmed_quantity': material.confirmed_quantity,
                              'dispatch_status': 'dispatched' if material.confirmed_quantity >= material.required_quantity else 'backordered',
                          }
                      )
                      if created:
                          messages.info(request, f"為申請單 {req.order_number} ({req.process_type}) 新增物料 {material.material_number}。")
                      else:
                          messages.info(request, f"更新申請單 {req.order_number} ({req.process_type}) 的物料 {material.material_number} 撥料數量。")
                      
                      affected_requisition_ids.add(req.pk)

                      # Check if all items for this requisition are fully dispatched
                      all_items_for_req = RequisitionItem.objects.filter(requisition=req)
                      all_dispatched = True
                      for item in all_items_for_req:
                          confirmed_qty = item.confirmed_quantity if item.confirmed_quantity is not None else Decimal('0')
                          if confirmed_qty < item.required_quantity:
                              all_dispatched = False
                              break
                      
                      if all_dispatched and req.status != 'dispatch_completed':
                          req.status = 'dispatch_completed'
                          req.dispatch_performed = True # Set dispatch_performed to True
                          req.save()
                          messages.success(request, f"申請單 {req.order_number} ({req.process_type}) 所有物料已撥料完成！")
                          redirect_to_requisition_pk = req.pk # Store the PK for redirection
                  # --- End New Logic ---

              except (ValueError, WorkOrderMaterial.DoesNotExist) as e:
                  print(f"  Error in loop: {e}") # DEBUG
                  messages.error(request, f"處理物料 ID  {key.split('_')[1]} 時發生錯誤: {e}，部分或所有變更可能未儲存。")
                  return redirect(f'{redirect_url}{query_string}')

      # --- New Logic: Handle WorkOrderMaterials not explicitly processed by user input ---
      for material in all_relevant_work_order_materials:
          # Check if a RequisitionItem already exists for this material and requisition
          # This covers both explicitly processed materials and previously existing items
          if not RequisitionItem.objects.filter(requisition=current_requisition, material_number=material.material_number).exists():
              # This material has no RequisitionItem yet, meaning it was not explicitly handled
              # and no RequisitionItem was created for it previously.
              # It should be considered 'backordered' as it was not dispatched.
              
              # Create RequisitionItem for this unprocessed material
              requisition_item, created = RequisitionItem.objects.update_or_create(
                  requisition=current_requisition,
                  material_number=material.material_number,
                  defaults={
                      'source_material': material,
                      'order_number': material.order_number,
                      'item_name': material.item_name,
                      'required_quantity': material.required_quantity,
                      'stock_quantity': Material.objects.filter(material_code=material.material_number).values_list('system_quantity', flat=True).first() or Decimal('0'),
                      'confirmed_quantity': Decimal('0'), # Set to 0 as it was not dispatched
                      'dispatch_status': 'backordered', # Explicitly set to backordered
                  }
              )
              if created:
                  messages.info(request, f"為申請單 {current_requisition.order_number} ({current_requisition.process_type}) 新增未撥物料 {material.material_number}。")
              else:
                  messages.info(request, f"更新申請單 {current_requisition.order_number} ({current_requisition.process_type}) 的未撥物料 {material.material_number} 狀態。")
              
              affected_requisition_ids.add(current_requisition.pk)
      # --- End New Logic ---

      # After processing all materials (both explicitly updated and implicitly backordered),
      # check the overall status of the requisition.
      # This part needs to be updated to use current_requisition instead of req from the loop.
      if current_requisition:
          all_items_for_req = RequisitionItem.objects.filter(requisition=current_requisition)
          all_dispatched = True
          for item in all_items_for_req:
              confirmed_qty = item.confirmed_quantity if item.confirmed_quantity is not None else Decimal('0')
              if confirmed_qty < item.required_quantity:
                  all_dispatched = False
                  break
          
          if all_dispatched and current_requisition.status != 'dispatch_completed':
              current_requisition.status = 'dispatch_completed'
              current_requisition.dispatch_performed = True # Set dispatch_performed to True
              current_requisition.save()
              messages.success(request, f"申請單 {current_requisition.order_number} ({current_requisition.process_type}) 所有物料已撥料完成！")
              redirect_to_requisition_pk = current_requisition.pk # Store the PK for redirection
          elif not all_dispatched and current_requisition.status == 'demand_submitted':
              # If some items are backordered, but none were dispatched yet, set to dispatch_in_progress
              current_requisition.status = 'dispatch_in_progress'
              current_requisition.save()
      
      # Send notifications for all affected requisitions
      for req_id in affected_requisition_ids:
          try:
              req_to_notify = Requisition.objects.get(pk=req_id)
              notify_requisition_shortages(req_to_notify)
          except Requisition.DoesNotExist:
              pass

      if updated_materials:
            print(f"Updated materials list: {updated_materials}") # DEBUG
            if updated_materials:
                messages.success(request, f"成功更新撥料數量: {','.join(updated_materials)}")
            
            if redirect_to_requisition_pk:
                # If a requisition completed dispatch, redirect to its detail page
                return redirect('requisitions:requisition_detail', pk=redirect_to_requisition_pk)
            else:
                # If no requisition completed dispatch, but materials were updated,
                # we need to find the requisition associated with the order_number and process_type
                # and redirect to its detail page.
                # If no such requisition exists, then redirect back to the work_order_material_list.
                
                # Try to find a relevant requisition based on the order_number and process_type_filter
                # This assumes that the order_number and process_type_filter are always present in the POST data
                # and correspond to a single requisition that was being worked on.
                if order_number and process_type_filter:
                    try:
                        # Find the ProcessType object by its ID
                        process_type_obj = ProcessType.objects.get(pk=process_type_filter)
                        # Find the Requisition based on order_number and process_type name
                        current_requisition = Requisition.objects.filter(
                            order_number=order_number,
                            process_type=process_type_obj.name
                        ).first()
                        if current_requisition:
                            return redirect('requisitions:requisition_detail', pk=current_requisition.pk)
                    except ProcessType.DoesNotExist:
                        pass # Fallback to default redirect

                # Fallback if no specific requisition detail can be determined
                return redirect(f'{redirect_url}{query_string}')
      else:
          # If no materials were updated, but there might be implicitly backordered items,
          # we should still redirect to the requisition detail page if current_requisition exists.
          if current_requisition:
              messages.info(request, "沒有物料數量被明確更新，但已處理未撥物料狀態。")
              return redirect('requisitions:requisition_detail', pk=current_requisition.pk)
          else:
              messages.info(request, "沒有物料數量被更新。")
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


@login_required
def sync_storage_bins(request):
    """從上傳的 Excel 匯入儲格資料，更新 Material.bin 和 RequisitionItem.storage_bin"""
    if request.method != 'POST':
        return redirect('requisitions:work_order_material_list')

    if not (request.user.is_superuser or request.user.groups.filter(name__in=['撥料人員', '申請人員']).exists()):
        messages.error(request, "您沒有權限執行此操作。")
        return redirect('requisitions:work_order_material_list')

    excel_file = request.FILES.get('storage_bin_file')
    if not excel_file:
        messages.error(request, "請選擇一個 Excel 檔案。")
        redirect_url = request.META.get('HTTP_REFERER', reverse('requisitions:work_order_material_list'))
        return redirect(redirect_url)

    try:
        df = pd.read_excel(excel_file, dtype={'物料': str, '儲格': str}, engine='openpyxl')
        df.columns = df.columns.str.strip()

        if '物料' not in df.columns or '儲格' not in df.columns:
            messages.error(request, "Excel 檔案必須包含「物料」和「儲格」欄位。")
            redirect_url = request.META.get('HTTP_REFERER', reverse('requisitions:work_order_material_list'))
            return redirect(redirect_url)

        # 建立物料 -> 儲格對應
        bin_map = {}
        for _, row in df.iterrows():
            material_code = str(row['物料']).strip()
            bin_value = str(row['儲格']).strip()
            if material_code and bin_value and bin_value != 'nan':
                bin_map[material_code] = bin_value

        material_updated = 0
        item_updated = 0

        with transaction.atomic():
            # 1. 更新 Material.bin
            for material_code, bin_value in bin_map.items():
                updated = Material.objects.filter(material_code=material_code).exclude(bin=bin_value).update(bin=bin_value)
                material_updated += updated

            # 2. 更新 RequisitionItem.storage_bin
            items = RequisitionItem.objects.all()
            for item in items:
                new_bin = bin_map.get(item.material_number, '')
                if new_bin and item.storage_bin != new_bin:
                    item.storage_bin = new_bin
                    item.save(update_fields=['storage_bin'])
                    item_updated += 1

        messages.success(request, f"儲格匯入完成！Excel 共 {len(bin_map)} 筆。更新 Material {material_updated} 筆、RequisitionItem {item_updated} 筆。")

    except Exception as e:
        messages.error(request, f"處理 Excel 時發生錯誤: {e}")

    redirect_url = request.META.get('HTTP_REFERER', reverse('requisitions:work_order_material_list'))
    return redirect(redirect_url)