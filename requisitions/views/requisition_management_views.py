from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.utils import timezone
from ..forms import RequisitionForm, UploadFileForm, RequisitionItemMaterialConfirmationFormSet, RequisitionItemSignOffFormSet
from ..models import Requisition, RequisitionItem, MaterialListVersion, WorkOrderMaterial, Inventory, MachineModel, ProcessType
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

def _filter_requisitions(request, sort_by='created_at', order='desc', material_status_filter=None, order_number_search=None):
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
    order_number_search = request.GET.get('order_number_search')
    material_status_selected = request.GET.get('material_status')

    # show_results should always be true, the template will handle empty results
    show_results = True 
    
    unique_requisitions = _filter_requisitions(request, 
                                                sort_by=model_sort_by, 
                                                order=order, 
                                                material_status_filter=material_status_selected,
                                                order_number_search=order_number_search)

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
        return redirect('requisition_list')

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
                requisition.save()
                messages.success(request, "撥料申請單建立成功！")
                return redirect('requisition_list')
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
        return redirect('homepage')

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

    # Get process types for the filter dropdown
    process_types = ProcessType.objects.exclude(name='nan').order_by('name') # Exclude 'nan' named process types
    process_type_choices = [(pt.name, pt.name) for pt in process_types] # Assuming process_type in Requisition is CharField storing name

    return render(request, 'requisitions/requisition_history.html', {
        'history_requisitions': requisitions_page,
        'process_type_choices': process_type_choices, # Add process type choices to context
    })

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
def requisition_sign_off(request, pk):
    requisition = get_object_or_404(Requisition, pk=pk)

    if not request.user.groups.filter(name='申請人員').exists() and not request.user.is_superuser:
        messages.error(request, "您沒有權限執行最終簽收操作。")
        return redirect('requisition_list')

    if request.method == 'POST':
        with transaction.atomic():
            for key in request.POST.keys():
                if key.startswith('signed_off_'):
                    material_id = key.split('_')[-1]
                    try:
                        material = WorkOrderMaterial.objects.get(id=material_id)
                        material.is_signed_off = True
                        material.save()
                    except WorkOrderMaterial.DoesNotExist:
                        messages.error(request, f"物料 ID {material_id} 不存在。")
                        continue
            
            # Check if all materials for this requisition are signed off
            all_materials = WorkOrderMaterial.objects.filter(
                order_number=requisition.order_number,
                process_type__name=requisition.process_type,
                is_active=True
            )
            if all(m.is_signed_off for m in all_materials):
                requisition.status = 'completed'
                requisition.sign_off_by = request.user
                requisition.sign_off_date = timezone.now()
                requisition.save()
                messages.success(request, "撥料單已全部最終簽收！")
            else:
                messages.success(request, "部分物料已簽收。")

    return redirect('requisition_list')


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