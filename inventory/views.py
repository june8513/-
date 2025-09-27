from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from .models import Material, Stocktake, StocktakeItem, MaterialTransaction
import pandas as pd
from django.db import transaction
from django.utils import timezone
from django.http import HttpResponse, JsonResponse
from django.db.models import Count, Q, F
import json

# View for importing master material data from Excel
@login_required
def import_material_master(request):
    if request.method == 'POST' and request.FILES.get('excel_file'):
        excel_file = request.FILES['excel_file']
        expected_columns = ['儲存地點', '儲格', '物料', '物料說明', '未限制']
        try:
            df = pd.read_excel(excel_file, dtype={'儲存地點': str, '儲格': str, '物料': str})

            if not all(col in df.columns for col in expected_columns):
                missing_cols = ", ".join([col for col in expected_columns if col not in df.columns])
                messages.error(request, f"Excel 檔案缺少必要的欄位，請檢查是否包含: {missing_cols}")
                return redirect('inventory_update')

            with transaction.atomic():
                for index, row in df.iterrows():
                    material_code = row['物料']
                    new_quantity = row['未限制']
                    
                    material, created = Material.objects.update_or_create(
                        material_code=material_code,
                        defaults={
                            'location': row['儲存地點'],
                            'bin': row['儲格'],
                            'material_description': row['物料說明'],
                            'system_quantity': new_quantity,
                            'latest_counted_quantity': new_quantity, # Sync counted quantity
                            'last_counted_by': request.user, # Attribute the change
                            'last_counted_date': timezone.now(), # Update the date
                        }
                    )
                    
                    MaterialTransaction.objects.create(
                        material=material,
                        user=request.user,
                        transaction_type='INITIAL_IMPORT' if created else 'MANUAL_UPDATE',
                        quantity_change=new_quantity, # This might need better logic
                        new_system_quantity=material.system_quantity,
                        notes=f"透過 Excel 檔案更新: {excel_file.name}"
                    )
            messages.success(request, f"成功匯入/更新 {len(df)} 筆主物料資料。")

        except Exception as e:
            messages.error(request, f"匯入失敗，發生預期外的錯誤: {e}")
            print(f"Error importing Excel: {e}")

    return redirect('inventory_home')

# View for listing master materials and providing import/stocktake creation options
@login_required
def material_list(request):
    materials = Material.objects.all()

    # Filtering for material list
    location_filter = request.GET.get('location_filter')
    if location_filter:
        materials = materials.filter(location__icontains=location_filter)

    # Sorting logic
    sort_by = request.GET.get('sort_by', 'material_code') # Default sort by material_code
    order = request.GET.get('order', 'asc') # Default order ascending

    # Validate sort_by to prevent SQL injection or invalid field names
    valid_sort_fields = ['location', 'bin', 'material_code', 'material_description', 'system_quantity', 'last_counted_date', 'latest_counted_quantity'] # Removed latest_difference
    if sort_by not in valid_sort_fields:
        sort_by = 'material_code' # Fallback to default

    if order == 'desc':
        materials = materials.order_by(f'-{sort_by}')
    else:
        materials = materials.order_by(sort_by)

    return render(request, 'inventory/material_list.html', {
        'materials': materials,
        'location_filter': location_filter,
        'sort_by': sort_by, # Pass to template
        'order': order # Pass to template
    })

@login_required
@transaction.atomic
def update_material_quantities(request):
    if request.method == 'POST':
        updated_materials = []
        for key, value in request.POST.items():
            if key.startswith('quantity_') and value:
                try:
                    material_id = int(key.split('_')[1])
                    quantity_change = int(value)

                    if quantity_change == 0:
                        continue

                    material = Material.objects.get(pk=material_id)
                    original_quantity = material.system_quantity
                    material.system_quantity += quantity_change
                    material.save()

                    transaction_type = 'RETURN' if quantity_change > 0 else 'ALLOCATION'

                    MaterialTransaction.objects.create(
                        material=material,
                        user=request.user,
                        transaction_type=transaction_type,
                        quantity_change=quantity_change,
                        new_system_quantity=material.system_quantity,
                        notes="手動庫存操作"
                    )
                    updated_materials.append(f"{material.material_code} ({quantity_change:+.0f})")

                except (ValueError, Material.DoesNotExist) as e:
                    messages.error(request, f"處理物料 ID {key.split('_')[1]} 時發生錯誤: {e}，部分或所有變更可能未儲存。")
                    # Since we are in a transaction, any error will cause a rollback.
                    return redirect('material_list')
        
        if updated_materials:
            messages.success(request, f"成功更新庫存: {', '.join(updated_materials)}")
        else:
            messages.info(request, "沒有偵測到任何庫存變動。")

    return redirect('material_list')

# View for creating a new stocktake from selected materials
@login_required
@transaction.atomic
def create_stocktake_from_selection(request):
    if request.method == 'POST':
        selected_material_ids = request.POST.getlist('selected_materials')
        stocktake_name = request.POST.get('stocktake_name') # Get the custom name

        if not selected_material_ids:
            messages.warning(request, "請選擇至少一個物料來建立盤點單。")
            return redirect('material_list')

        materials_to_stocktake = Material.objects.filter(id__in=selected_material_ids)

        # Create a new Stocktake header
        stocktake = Stocktake.objects.create(
            created_by=request.user,
            stocktake_id=f"ST-{Stocktake.objects.count() + 1}",
            name=stocktake_name if stocktake_name else None # Assign custom name
        )

        # Create StocktakeItems for each selected material
        for material in materials_to_stocktake:
            StocktakeItem.objects.create(
                stocktake=stocktake,
                material=material,
                system_quantity_on_record=material.system_quantity # Snapshot current system quantity
            )
        messages.success(request, f"成功建立盤點單 {stocktake.stocktake_id}。")
        return redirect('stocktake_detail', pk=stocktake.pk)
    return redirect('material_list')

# View for listing all stocktakes
@login_required
def stocktake_list(request):
    stocktakes = Stocktake.objects.all().order_by('-created_at')
    return render(request, 'inventory/stocktake_list.html', {'stocktakes': stocktakes})

# View for stocktake detail and counting
@login_required
def stocktake_detail(request, pk):
    stocktake = get_object_or_404(Stocktake, pk=pk)
    items = stocktake.items.all()

    # Filtering for stocktake items
    location_filter = request.GET.get('location_filter')
    if location_filter:
        items = items.filter(material__location__icontains=location_filter)

    # Sorting logic for StocktakeItem
    sort_by = request.GET.get('sort_by', 'material__material_code') # Default sort
    order = request.GET.get('order', 'asc') # Default order ascending

    # Valid sort fields for StocktakeItem (using __ for related fields)
    valid_sort_fields = [
        'material__location', 'material__bin', 'material__material_code',
        'material__material_description', 'system_quantity_on_record',
        'counted_quantity', 'status'
    ]
    if sort_by not in valid_sort_fields:
        sort_by = 'material__material_code' # Fallback to default

    if order == 'desc':
        items = items.order_by(f'-{sort_by}')
    else:
        items = items.order_by(sort_by)

    return render(request, 'inventory/stocktake_detail.html', {
        'stocktake': stocktake, 
        'items': items,
        'location_filter': location_filter,
        'sort_by': sort_by, # Pass to template
        'order': order # Pass to template
    })

# View to handle saving quantities and completing stocktake
@login_required
@transaction.atomic
def handle_stocktake_actions(request, pk):
    if request.method == 'POST':
        stocktake = get_object_or_404(Stocktake, pk=pk)
        action = request.POST.get('action')
        
        # --- DEBUGGING: Display request.POST content --- START
        # messages.info(request, f"Received POST data: {request.POST}") # Removed debugging line
        # --- DEBUGGING: Display request.POST content --- END

        if stocktake.status == '進行中':
            # Removed notes saving logic

            # Always save quantities (for save_quantities or complete_stocktake actions)
            for item in stocktake.items.all():
                counted_quantity_str = request.POST.get(f'counted_quantity_{item.id}')
                if counted_quantity_str is not None and counted_quantity_str != '':
                    try:
                        item.counted_quantity = int(counted_quantity_str)
                        item.status = '已盤點'
                        item.save()
                    except ValueError:
                        messages.error(request, f"物料 {item.material.material_code} 的盤點數量無效，請輸入數字。")
                        return redirect('stocktake_detail', pk=pk)
            
            if action == 'save_quantities':
                messages.success(request, "盤點資料已更新。")
            elif action == 'complete_stocktake':
                # Then, mark the stocktake as complete
                stocktake.status = '已完成'
                # stocktake.save() # This save will be done at the end

                # Update last_counted_date and latest_counted_quantity for materials that were counted
                for item in stocktake.items.filter(status='已盤點'):
                    material = item.material
                    material.last_counted_date = timezone.now()
                    material.latest_counted_quantity = item.counted_quantity
                    # latest_difference is now a property, no need to save it
                    material.save()

                messages.success(request, f"盤點單 {stocktake.stocktake_id} 已標記為完成，並已儲存盤點數量。")
        else:
            messages.warning(request, "此盤點單已完成，無法再修改。")
    
    # Save the stocktake object (notes and status) once at the end for save_quantities or complete_stocktake
    # This save is only reached if action is not 'save_notes'
    stocktake.save()
    return redirect('stocktake_detail', pk=pk)

# View for exporting stocktake differences
@login_required
def export_stocktake_differences(request, pk):
    stocktake = get_object_or_404(Stocktake, pk=pk)
    
    # Get all items for this stocktake
    all_items = stocktake.items.all()

    # Filter for items with differences
    # Since 'difference' is a property, we filter in Python
    items_with_differences = [item for item in all_items if item.difference is not None and item.difference != 0]

    # Prepare data for DataFrame
    data = []
    for item in items_with_differences:
        data.append({
            '盤點單號': stocktake.stocktake_id,
            '盤點單名稱': stocktake.name if stocktake.name else stocktake.stocktake_id,
            '庫位': item.material.location,
            '儲格': item.material.bin,
            '物料': item.material.material_code,
            '物料說明': item.material.material_description,
            '系統數量': item.system_quantity_on_record,
            '盤點數量': item.counted_quantity if item.counted_quantity is not None else '',
            '差異': item.difference if item.difference is not None else '',
            '狀態': item.status,
        })

    df = pd.DataFrame(data)

    # Create Excel response
    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = f'attachment; filename="{stocktake.stocktake_id}_differences.xlsx"'
    df.to_excel(response, index=False, engine='openpyxl')
    return response

# View for exporting master material differences
@login_required
def export_master_material_differences(request):
    # Get all materials
    all_materials = Material.objects.all()

    # Filter for materials with differences OR not counted
    materials_to_export = [
        m for m in all_materials 
        if (m.current_difference is not None and m.current_difference != 0) or m.latest_counted_quantity is None
    ]

    # Prepare data for DataFrame
    data = []
    for material in materials_to_export:
        data.append({
            '庫位': material.location,
            '儲格': material.bin,
            '物料': material.material_code,
            '物料說明': material.material_description,
            '系統庫存數量': material.system_quantity,
            '最新盤點數量': material.latest_counted_quantity if material.latest_counted_quantity is not None else '',
            '最新差異': material.current_difference if material.current_difference is not None else '',
            '上次盤點日期': material.last_counted_date.strftime("%Y-%m-%d %H:%M") if material.last_counted_date else '',
        })

    df = pd.DataFrame(data)

    # Create Excel response
    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = f'attachment; filename="master_material_report.xlsx"' # Changed filename
    df.to_excel(response, index=False, engine='openpyxl')
    return response

@login_required
def inventory_dashboard(request):
    if not request.user.is_superuser:
        messages.error(request, "您沒有權限訪問此頁面。")
        return redirect('homepage')
    
    return render(request, 'inventory/inventory_dashboard.html')

@login_required
def inventory_update_view(request):
    if not request.user.is_superuser:
        messages.error(request, "您沒有權限訪問此頁面。")
        return redirect('inventory_home')
    return render(request, 'inventory/inventory_update.html')

@login_required
def stocktake_location_list(request):
    if not request.user.is_superuser:
        messages.error(request, "您沒有權限訪問此頁面。")
        return redirect('inventory_home')

    locations_stats = Material.objects.exclude(
        Q(location__isnull=True) | Q(location__exact='')
    ).values('location').annotate(
        total_items=Count('id'),
        uncounted_items=Count('id', filter=Q(latest_counted_quantity__isnull=True))
    ).order_by('location')

    context = {
        'locations_stats': locations_stats
    }
    return render(request, 'inventory/stocktake_location_list.html', context)

@login_required
def stocktake_detail_by_location(request, location_name):
    if not request.user.is_superuser:
        messages.error(request, "您沒有權限訪問此頁面。")
        return redirect('inventory_home')

    materials = Material.objects.filter(location=location_name).order_by('bin')
    
    context = {
        'location_name': location_name,
        'materials': materials,
    }
    return render(request, 'inventory/stocktake_detail.html', context)

def update_counted_quantity(request):
    if not request.user.is_authenticated:
        return JsonResponse({'status': 'error', 'message': '用戶未登入或會話已過期，請重新登入。'}, status=401)

    if not request.user.is_superuser:
        return JsonResponse({'status': 'error', 'message': '權限不足'}, status=403)

    if request.method == 'POST':
        try:
            data = json.loads(request.body)
            material_id = data.get('material_id')
            quantity = data.get('quantity')

            material = Material.objects.get(pk=material_id)
            
            if quantity == '' or quantity is None:
                material.latest_counted_quantity = None
            else:
                material.latest_counted_quantity = int(quantity)

            material.last_counted_by = request.user
            material.last_counted_date = timezone.now()
            material.save()
            
            return JsonResponse({
                'status': 'success',
                'message': '數量已更新',
                'last_counted_by': request.user.username,
                'last_counted_date': material.last_counted_date.isoformat()
            })
        except Exception as e:
            return JsonResponse({'status': 'error', 'message': str(e)}, status=400)
    
    return JsonResponse({'status': 'error', 'message': '無效的請求'}, status=400)

@login_required
def difference_location_list(request):
    if not request.user.is_superuser:
        messages.error(request, "您沒有權限訪問此頁面。")
        return redirect('inventory_home')

    locations_stats = Material.objects.exclude(
        Q(location__isnull=True) | Q(location__exact='')
    ).exclude(
        latest_counted_quantity__isnull=True
    ).exclude(
        system_quantity=F('latest_counted_quantity')
    ).values('location').annotate(
        difference_items=Count('id')
    ).order_by('location')

    context = {
        'locations_stats': locations_stats
    }
    return render(request, 'inventory/difference_location_list.html', context)

@login_required
def difference_detail_by_location(request, location_name):
    if not request.user.is_superuser:
        messages.error(request, "您沒有權限訪問此頁面。")
        return redirect('inventory_home')

    materials = Material.objects.filter(
        location=location_name
    ).exclude(
        latest_counted_quantity__isnull=True
    ).exclude(
        system_quantity=F('latest_counted_quantity')
    ).order_by('bin')

    # Calculate difference for each material
    for mat in materials:
        mat.difference = mat.latest_counted_quantity - mat.system_quantity

    context = {
        'location_name': location_name,
        'materials': materials,
    }
    return render(request, 'inventory/difference_detail.html', context)

@login_required
def export_differences_excel(request, location_name):
    if not request.user.is_superuser:
        messages.error(request, "您沒有權限執行此操作。")
        return redirect('inventory_home')

    materials = Material.objects.filter(
        location=location_name
    ).exclude(
        latest_counted_quantity__isnull=True
    ).exclude(
        system_quantity=F('latest_counted_quantity')
    ).order_by('bin')

    data = []
    for material in materials:
        difference = material.latest_counted_quantity - material.system_quantity
        data.append({
            '儲格': material.bin,
            '物料': material.material_code,
            '物料說明': material.material_description,
            '庫存數': material.system_quantity,
            '盤點數': material.latest_counted_quantity,
            '差異數': difference,
            '盤點人員': material.last_counted_by.username if material.last_counted_by else '',
            '盤點時間': timezone.localtime(material.last_counted_date).strftime('%Y-%m-%d %H:%M') if material.last_counted_date else '',
        })

    df = pd.DataFrame(data)
    
    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = f'attachment; filename="盤點差異報告_{location_name}_{timezone.now().strftime("%Y%m%d")}.xlsx"'
    df.to_excel(response, index=False, engine='openpyxl')
    
    return response