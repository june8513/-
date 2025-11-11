from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.utils import timezone
from ..forms import RequisitionForm, UploadFileForm, RequisitionItemMaterialConfirmationFormSet, RequisitionItemSignOffFormSet
from ..models import Requisition, RequisitionItem, WorkOrderMaterial, Inventory, MachineModel, ProcessType
from inventory.models import Material
from django.db import transaction
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
import json
from decimal import Decimal, InvalidOperation
import datetime

def _filter_requisitions(request, sort_by='created_at', order='desc', material_status_filter=None, order_number_search=None, status_filter_list=None):
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
    base_queryset = Requisition.objects.filter(dispatch_performed=False, is_archived=False).order_by(order_field).select_related('applicant')
    if status_filter_list:
        base_queryset = base_queryset.filter(status__in=status_filter_list)

    # Add search by order number
    if order_number_search:
        base_queryset = base_queryset.filter(order_number__icontains=order_number_search)

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
                items__isnull=False # Check if there are any RequisitionItems directly linked
            ).distinct()
        elif material_status_filter == 'no_materials':
            all_requisitions = all_requisitions.filter(
                items__isnull=True # Check if there are no RequisitionItems directly linked
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
    order_number_search = request.GET.get('order_number_search')
    material_status_selected = request.GET.get('material_status')

    # show_results should always be true, the template will handle empty results
    show_results = True 
    
    unique_requisitions = _filter_requisitions(request, 
                                                sort_by=model_sort_by, 
                                                order=order, 
                                                material_status_filter=material_status_selected,
                                                order_number_search=order_number_search,
                                                status_filter_list=['demand_submitted', 'dispatch_in_progress'])

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
        'process_type_choices': process_type_choices,
        'material_status_choices': material_status_choices, # New choices for template
        'selected_process_type': process_type_selected,
        'selected_material_status': material_status_selected, # Pass selected material status
        'sort_by': sort_by,
        'order': order,
        'query_params': request.GET.urlencode(),
        'show_results': show_results,
    }
    return render(request, 'requisitions/requisition_list.html', context)

@login_required
def requisition_detail(request, pk):
    requisition = get_object_or_404(Requisition, pk=pk)

    # Permissions check: Only applicant, material handler, or admin can view
    is_admin = request.user.is_superuser
    is_applicant = request.user == requisition.applicant
    is_material_handler = request.user.groups.filter(name='撥料人員').exists()

    # --- DEBUG LOGS ---
    print(f"DEBUG: requisition_detail view for Requisition PK: {pk}")
    print(f"DEBUG: Current User: {request.user.username} (ID: {request.user.id})")
    print(f"DEBUG: Is Admin: {is_admin}")
    print(f"DEBUG: Is Applicant (for this requisition): {is_applicant}")
    print(f"DEBUG: Is Material Handler: {is_material_handler}")
    print(f"DEBUG: Requisition Applicant: {requisition.applicant.username} (ID: {requisition.applicant.id})")
    # --- END DEBUG LOGS ---

    if not (is_admin or is_applicant or is_material_handler):
        messages.error(request, "您沒有權限查看此撥料申請單詳情。")
        return redirect('requisitions:requisition_list')

    # Fetch all RequisitionItems for this requisition
    all_items = RequisitionItem.objects.filter(requisition=requisition).order_by('material_number')

    # Categorize items
    dispatched_items = all_items.filter(dispatch_status='dispatched', is_signed_off=False)
    backordered_items = all_items.filter(dispatch_status='backordered', is_signed_off=False)
    signed_off_items = all_items.filter(is_signed_off=True)

    # --- DEBUG LOGS for Requisition Items ---
    print(f"DEBUG: Requisition {pk} - Total RequisitionItems: {all_items.count()}")
    print(f"DEBUG: Requisition {pk} - Dispatched Items (not signed off): {dispatched_items.count()}")
    for item in dispatched_items:
        print(f"  - Dispatched Item PK: {item.pk}, Material: {item.material_number}, Confirmed Qty: {item.confirmed_quantity}, Required Qty: {item.required_quantity}, Dispatch Status: {item.dispatch_status}, Signed Off: {item.is_signed_off}")
    print(f"DEBUG: Requisition {pk} - Backordered Items (not signed off): {backordered_items.count()}")
    for item in backordered_items:
        print(f"  - Backordered Item PK: {item.pk}, Material: {item.material_number}, Confirmed Qty: {item.confirmed_quantity}, Required Qty: {item.required_quantity}, Dispatch Status: {item.dispatch_status}, Signed Off: {item.is_signed_off}")
    print(f"DEBUG: Requisition {pk} - Signed Off Items: {signed_off_items.count()}")
    for item in signed_off_items:
        print(f"  - Signed Off Item PK: {item.pk}, Material: {item.material_number}, Confirmed Qty: {item.confirmed_quantity}, Required Qty: {item.required_quantity}, Dispatch Status: {item.dispatch_status}, Signed Off: {item.is_signed_off}")
    # --- END DEBUG LOGS ---

    context = {
        'requisition': requisition,
        'dispatched_items': dispatched_items,
        'backordered_items': backordered_items,
        'signed_off_items': signed_off_items,
        'is_admin': is_admin,
        'is_applicant': is_applicant,
        'is_material_handler': is_material_handler,
    }
    return render(request, 'requisitions/requisition_detail.html', context)

@login_required
def archived_requisition_list(request):
    is_admin = request.user.is_superuser
    is_applicant = request.user.groups.filter(name='申請人員').exists()
    is_material_handler = request.user.groups.filter(name='撥料人員').exists()

    if not is_admin and not is_applicant and not is_material_handler:
        messages.error(request, "您沒有權限查看此頁面。")
        return redirect('homepage')

    sort_by = request.GET.get('sort_by', 'created_at')
    order = request.GET.get('order', 'desc')

    sort_mapping = {
        'work_order_number': 'order_number',
        'applicant': 'applicant__username',
        'request_date': 'request_date',
        'process_type': 'process_type',
        'status': 'status',
        'created_at': 'created_at',
    }
    model_sort_by = sort_mapping.get(sort_by, 'created_at')

    order_field = model_sort_by
    if order == 'desc':
        order_field = '-' + order_field

    base_queryset = Requisition.objects.filter(is_archived=True).order_by(order_field).select_related('applicant')

    if is_admin:
        all_requisitions = base_queryset
    elif is_applicant:
        all_requisitions = base_queryset.filter(applicant=request.user)
    elif is_material_handler:
        all_requisitions = base_queryset
    else:
        all_requisitions = Requisition.objects.none()

    process_type_filter = request.GET.get('process_type')
    if process_type_filter:
        all_requisitions = all_requisitions.filter(process_type=process_type_filter)

    unique_requisitions = []
    seen_combinations = set()
    for req in all_requisitions: 
        combination = (req.order_number, req.process_type)
        if combination not in seen_combinations:
            machine_model_names = list(WorkOrderMaterial.objects.filter(
                order_number=req.order_number
            ).values_list('machine_model__name', flat=True).distinct().order_by('machine_model__name'))
            
            req.machine_models_display = ", ".join(machine_model_names)
            unique_requisitions.append(req)
            seen_combinations.add(combination)
    
    paginator = Paginator(unique_requisitions, 10)
    page = request.GET.get('page')
    try:
        requisitions_page = paginator.page(page)
    except PageNotAnInteger:
        requisitions_page = paginator.page(1)
    except EmptyPage:
        requisitions_page = paginator.page(paginator.num_pages)

    process_types = Requisition.objects.order_by('process_type').values_list('process_type', flat=True).distinct()
    process_type_choices = [(pt, pt) for pt in process_types if pt]

    context = {
        'requisitions': requisitions_page,
        'is_admin': is_admin,
        'is_applicant': is_applicant,
        'is_material_handler': is_material_handler,
        'process_type_choices': process_type_choices,
        'selected_process_type': process_type_filter,
        'sort_by': sort_by,
        'order': order,
        'query_params': request.GET.urlencode(),
    }
    return render(request, 'requisitions/archived_requisition_list.html', context)

@login_required
def requisition_create(request):
    if not request.user.groups.filter(name='申請人員').exists() and not request.user.is_superuser:
        messages.error(request, "您沒有權限建立撥料申請單。")
        return redirect('requisitions:requisition_list')

    if request.method == 'POST':
        order_number = request.POST.get('order_number') # Get order_number from POST data

        # Re-generate choices for process_type based on the submitted order_number
        material_process_type_ids = WorkOrderMaterial.objects.filter(
            order_number=order_number,
            is_active=True # Only consider active materials
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
        if form.is_valid(): # This is where validation happens
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
                requisition.status = 'demand_submitted' # Set the new status
                requisition.save()
                messages.success(request, "撥料申請單建立成功，等待撥料人員處理！")
                return redirect('requisitions:requisition_list')
            except IntegrityError:
                messages.error(request, "此訂單單號在該需求流程中已存在，請使用不同的訂單單號或需求流程。")
            except Exception as e:
                messages.error(request, f"建立撥料申請單時發生錯誤: {e}")
                import traceback
                print(traceback.format_exc())
        else: # Form is not valid
            pass # No debug print needed
    else:
        form = RequisitionForm()
    return render(request, 'requisitions/requisition_create.html', {'form': form})


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
def requisition_history(request):
    is_admin = request.user.is_superuser
    is_applicant = request.user.groups.filter(name='申請人員').exists()
    is_material_handler = request.user.groups.filter(name='撥料人員').exists()

    if not is_admin and not is_applicant and not is_material_handler:
        messages.error(request, "您沒有權限查看此頁面。")
        return redirect('requisitions:homepage')

    history_requisitions_qs = Requisition.objects.filter(dispatch_performed=True, is_archived=False).select_related('applicant').order_by('-updated_at')

    work_order_number = request.GET.get('order_number') # Corrected key
    applicant_username = request.GET.get('applicant_username')
    process_type = request.GET.get('process_type')
    start_date = request.GET.get('start_date')
    end_date = request.GET.get('end_date')
    material_or_item_search = request.GET.get('material_or_item_search')
    status_filter = request.GET.get('status') # Get the status filter

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
    if status_filter: # Apply the status filter
        history_requisitions_qs = history_requisitions_qs.filter(status=status_filter)
    
    if material_or_item_search:
        matching_items_subquery = RequisitionItem.objects.filter(
            Q(material_number__icontains=material_or_item_search) |
            Q(item_name__icontains=material_or_item_search),
            requisition=OuterRef('pk') # Directly link to Requisition
        )
        history_requisitions_qs = history_requisitions_qs.filter(
            Exists(matching_items_subquery)
        ).distinct()

    paginator = Paginator(history_requisitions_qs, 10)
    page_number = request.GET.get('page')
    requisitions_page = paginator.get_page(page_number)

    for req in requisitions_page:
        machine_model_names = list(WorkOrderMaterial.objects.filter(
            order_number=req.order_number
        ).values_list('machine_model__name', flat=True).distinct().order_by('machine_model__name'))
        req.machine_models_display = ", ".join(machine_model_names)

    # Get process types for the filter dropdown
    process_types = ProcessType.objects.exclude(name='nan').order_by('name') # Exclude 'nan' named process types
    process_type_choices = [(pt.name, pt.name) for pt in process_types] # Assuming process_type in Requisition is CharField storing name

    return render(request, 'requisitions/requisition_history.html', {
        'history_requisitions': requisitions_page,
        'process_type_choices': process_type_choices, # Add process type choices to context
    })



@login_required
def requisition_sign_off(request, pk):
    requisition = get_object_or_404(Requisition, pk=pk)

    is_admin = request.user.is_superuser
    is_applicant = request.user == requisition.applicant

    if not (is_applicant or is_admin):
        messages.error(request, "您沒有權限執行最終簽收操作。")
        return redirect('requisitions:requisition_list')

    if request.method == 'POST':
        with transaction.atomic():
            signed_off_count = 0
            if 'confirm_all_sign_off' in request.POST:
                # Sign off all dispatched items for this requisition
                dispatched_items_to_sign_off = RequisitionItem.objects.filter(
                    requisition=requisition,
                    dispatch_status='dispatched',
                    is_signed_off=False
                )
                for item in dispatched_items_to_sign_off:
                    item.is_signed_off = True
                    item.sign_off_by = request.user
                    item.sign_off_date = timezone.now()
                    item.save()
                    signed_off_count += 1
                if signed_off_count > 0:
                    messages.success(request, f"成功簽收 {signed_off_count} 筆已撥料項目。")
                else:
                    messages.info(request, "沒有新的已撥料項目需要簽收。")
            else: # Process individual sign-off buttons
                for key in request.POST.keys():
                    if key.startswith('sign_off_item_'):
                        item_pk = key.split('_')[-1]
                        try:
                            item = RequisitionItem.objects.get(pk=item_pk, requisition=requisition)
                            if not item.is_signed_off:
                                item.is_signed_off = True
                                item.sign_off_by = request.user
                                item.sign_off_date = timezone.now()
                                item.save()
                                signed_off_count += 1
                        except RequisitionItem.DoesNotExist:
                            messages.error(request, f"物料項目 ID {item_pk} 不存在。")
                            continue
                if signed_off_count > 0:
                    messages.success(request, f"成功簽收 {signed_off_count} 筆物料項目。")

            # Check if all relevant RequisitionItems for this requisition are signed off
            # We consider items that were dispatched or backordered
            all_relevant_items = RequisitionItem.objects.filter(
                requisition=requisition,
                dispatch_status__in=['dispatched', 'backordered']
            )
            if all_relevant_items.exists() and all(item.is_signed_off for item in all_relevant_items):
                requisition.status = 'signed_off'
                requisition.sign_off_by = request.user
                requisition.sign_off_date = timezone.now()
                requisition.save()
                messages.success(request, "撥料單已全部最終簽收！")
            elif all_relevant_items.exists():
                messages.info(request, "部分物料已簽收，但仍有未簽收項目。")
            else:
                messages.warning(request, "此申請單沒有可簽收的物料項目。")

    return redirect('requisitions:requisition_detail', pk=requisition.pk)




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
    
    # This function is incomplete in the original views.py, I will just move it as is.
    # It needs to return JSON data for requisition items.
    # For now, I'll return an empty JSON response or a placeholder.
    return JsonResponse({'items': []})





@login_required
def sign_off_item(request, pk, item_pk): # Removed version_pk
    requisition = get_object_or_404(Requisition, pk=pk)
    item = get_object_or_404(RequisitionItem, pk=item_pk, requisition=requisition) # Directly link to Requisition

    is_admin = request.user.is_superuser
    is_applicant = request.user == requisition.applicant # Consistent with requisition_sign_off

    if not is_applicant and not is_admin:
        return JsonResponse({'success': False, 'message': "您沒有權限簽收物料項目。"}, status=403)

    if request.method == 'POST':
        if not item.is_signed_off:
            item.is_signed_off = True
            item.save()