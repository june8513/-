from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from requisitions.models import Requisition, WorkOrderMaterial, Inventory
from django.db import transaction
import json
from decimal import Decimal
from django.http import JsonResponse, HttpResponse
from django.db.models import F, ExpressionWrapper, DecimalField, Subquery, OuterRef
from django.db.models.functions import Coalesce

@login_required
def finished_goods_dispatch(request):
    return render(request, 'requisitions/finished_goods_dispatch.html')

@login_required
def update_dispatch_note(request, pk):
    requisition = get_object_or_404(Requisition, pk=pk)

    is_admin = request.user.is_superuser
    is_applicant = request.user == requisition.applicant
    is_material_handler = request.user.groups.filter(name='撥料人員').exists()

    if not (is_admin or is_applicant or is_material_handler):
        messages.error(request, "您沒有權限訪問此頁面。")
        return redirect('core:homepage') # Changed from 'homepage'
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
    requisition = get_object_or_404(Requisition, pk=pk)

    is_admin = request.user.is_superuser
    is_applicant = request.user == requisition.applicant
    is_material_handler = request.user.groups.filter(name='撥料人員').exists()

    if not (is_admin or is_applicant or is_material_handler):
        return JsonResponse({'success': False, 'message': '權限不足'}, status=403)

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

@login_required
def generate_backorder_note(request, pk):
    requisition = get_object_or_404(Requisition, pk=pk)

    is_admin = request.user.is_superuser
    is_applicant = request.user == requisition.applicant
    is_material_handler = request.user.groups.filter(name='撥料人員').exists()

    if not (is_admin or is_applicant or is_material_handler):
        messages.error(request, "您沒有權限訪問此頁面。")
        return redirect('core:homepage') # Changed from 'requisitions:homepage'

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
def supplement_material(request):
    # Placeholder function
    return HttpResponse("This is a placeholder for supplement_material.")

@login_required
def dispatch_preparation_list(request):
    is_admin = request.user.is_superuser
    is_material_handler = request.user.groups.filter(name='撥料人員').exists()

    if not (is_admin or is_material_handler):
        messages.error(request, "您沒有權限訪問此頁面。")
        return redirect('core:homepage')

    # Filter requisitions that are awaiting dispatch and not archived
    requisitions_qs = Requisition.objects.filter(
        status='awaiting_dispatch',
        is_archived=False
    ).order_by('-created_at').select_related('applicant')

    # Add search/filter options if needed (e.g., by order_number, process_type)
    order_number_search = request.GET.get('order_number_search')
    if order_number_search:
        requisitions_qs = requisitions_qs.filter(order_number__icontains=order_number_search)

    process_type_filter = request.GET.get('process_type')
    if process_type_filter:
        requisitions_qs = requisitions_qs.filter(process_type=process_type_filter)

    # Pagination
    paginator = Paginator(requisitions_qs, 10) # Show 10 requisitions per page
    page_number = request.GET.get('page')
    requisitions_page = paginator.get_page(page_number)

    # Get unique process types for filter dropdown
    process_types = Requisition.objects.filter(
        status='awaiting_dispatch',
        is_archived=False
    ).values_list('process_type', flat=True).distinct().order_by('process_type')
    process_type_choices = [(pt, pt) for pt in process_types if pt]

    context = {
        'requisitions': requisitions_page,
        'process_type_choices': process_type_choices,
        'selected_process_type': process_type_filter,
        'order_number_search': order_number_search,
    }
    return render(request, 'requisitions/dispatch_preparation_list.html', context)
