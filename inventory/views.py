from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from .models import Material, Stocktake, StocktakeItem, MaterialTransaction, StorageLocation
import pandas as pd
from django.db import transaction
from django.utils import timezone
from django.http import HttpResponse, JsonResponse
from django.db.models import Count, Q, F, ExpressionWrapper, fields
from django.db import models
import json

# View for importing master material data from Excel
@login_required
def import_material_master(request):
    if request.method != 'POST' or not request.FILES.get('excel_file'):
        return redirect('inventory_update')

    selected_location_names = request.POST.getlist('selected_locations')
    new_locations_str = request.POST.get('new_locations', '')

    if new_locations_str:
        new_locations = [loc.strip() for loc in new_locations_str.split(',') if loc.strip()]
        selected_location_names.extend(new_locations)
    
    sync_location_names = sorted(list(set(selected_location_names)))

    if not sync_location_names:
        messages.error(request, "請至少選擇或輸入一個要同步的儲存地點。")
        return redirect('inventory_update')

    excel_file = request.FILES['excel_file']
    expected_columns = ['儲存地點', '儲格', '物料', '物料說明', '未限制']

    try:
        df = pd.read_excel(excel_file, dtype=str)
        df.columns = [col.strip() for col in df.columns]

        if not all(col in df.columns for col in expected_columns):
            missing_cols = ", ".join([col for col in expected_columns if col not in df.columns])
            messages.error(request, f"Excel 檔案缺少必要的欄位: {missing_cols}")
            return redirect('inventory_update')

        df_filtered = df[df['儲存地點'].isin(sync_location_names)].copy()
        df_filtered.dropna(subset=['物料'], inplace=True)
        df_filtered['未限制'] = pd.to_numeric(df_filtered['未限制'], errors='coerce').fillna(0).astype(int)

        duplicates = df_filtered[df_filtered.duplicated(subset=['物料'], keep=False)]
        if not duplicates.empty:
            duplicate_codes = ", ".join(duplicates['物料'].unique())
            messages.error(request, f"匯入失敗：在您選擇的儲存地點範圍內，Excel 檔案中包含重複的物料號碼: {duplicate_codes}")
            return redirect('inventory_update')

        with transaction.atomic():
            for loc_name in sync_location_names:
                StorageLocation.objects.get_or_create(name=loc_name)

            all_db_materials = {mat.material_code: mat for mat in Material.objects.all()}
            all_db_codes = set(all_db_materials.keys())
            scoped_db_codes = set(Material.objects.filter(location__name__in=sync_location_names).values_list('material_code', flat=True))
            excel_codes = set(df_filtered['物料'])

            codes_to_delete = scoped_db_codes - excel_codes
            codes_to_create = excel_codes - all_db_codes
            codes_to_update = excel_codes.intersection(all_db_codes)

            deleted_count = 0
            if codes_to_delete:
                materials_to_delete = Material.objects.filter(location__name__in=sync_location_names, material_code__in=codes_to_delete)
                deleted_count = materials_to_delete.delete()[0]

            created_count = 0
            updated_count = 0
            df_to_process = df_filtered.set_index('物料')

            for code in codes_to_create:
                row = df_to_process.loc[code]
                loc_obj = StorageLocation.objects.get(name=row['儲存地點'])
                Material.objects.create(
                    material_code=code, location=loc_obj, bin=row['儲格'],
                    material_description=row['物料說明'], system_quantity=int(row['未限制']),
                    latest_counted_quantity=None
                )
                created_count += 1
            
            for code in codes_to_update:
                row = df_to_process.loc[code]
                mat_to_update = all_db_materials[code]
                new_system_quantity = int(row['未限制'])
                loc_obj = StorageLocation.objects.get(name=row['儲存地點'])
                
                if (mat_to_update.location != loc_obj or mat_to_update.bin != row['儲格'] or
                    mat_to_update.material_description != row['物料說明'] or mat_to_update.system_quantity != new_system_quantity):
                    
                    mat_to_update.location = loc_obj
                    mat_to_update.bin = row['儲格']
                    mat_to_update.material_description = row['物料說明']
                    mat_to_update.system_quantity = new_system_quantity
                    mat_to_update.save()
                    updated_count += 1

            summary_message = f"在庫存地點 [{', '.join(sync_location_names)}] 中同步完成。新增 {created_count} 筆，更新 {updated_count} 筆，刪除 {deleted_count} 筆物料。"
            messages.success(request, summary_message)

    except Exception as e:
        messages.error(request, f"匯入失敗，發生預期外的錯誤: {e}")

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
    
    # Get all distinct locations to populate the form
    locations = StorageLocation.objects.all().order_by('name')
    
    context = {
        'locations': locations
    }
    return render(request, 'inventory/inventory_update.html', context)

@login_required
def stocktake_location_list(request):
    if not request.user.is_superuser:
        messages.error(request, "您沒有權限訪問此頁面。")
        return redirect('inventory_home')

    locations_stats = StorageLocation.objects.annotate(
        total_items=Count('material'),
        uncounted_items=Count('material', filter=Q(material__latest_counted_quantity__isnull=True))
    ).order_by('name')

    context = {
        'locations_stats': locations_stats
    }
    return render(request, 'inventory/stocktake_location_list.html', context)

@login_required
def stocktake_detail_by_location(request, location_name):
    if not request.user.is_superuser:
        messages.error(request, "您沒有權限訪問此頁面。")
        return redirect('inventory_home')

    sort_by = request.GET.get('sort_by', 'bin') # Default sort by bin
    order = request.GET.get('order', 'asc')

    valid_sort_fields = ['bin', 'material_code', 'material_description', 'system_quantity', 'latest_counted_quantity', 'last_counted_by__username', 'last_counted_date']
    if sort_by not in valid_sort_fields:
        sort_by = 'bin'

    order_prefix = '-' if order == 'desc' else ''
    materials = Material.objects.filter(location__name=location_name).order_by(f'{order_prefix}{sort_by}')
    
    context = {
        'location_name': location_name,
        'materials': materials,
        'sort_by': sort_by,
        'order': order,
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

    locations_stats = StorageLocation.objects.annotate(
        difference_items=Count('material', 
            filter=~Q(material__system_quantity=F('material__latest_counted_quantity')) & 
                   Q(material__latest_counted_quantity__isnull=False)
        )
    ).filter(difference_items__gt=0).order_by('name')

    context = {
        'locations_stats': locations_stats
    }
    return render(request, 'inventory/difference_location_list.html', context)

@login_required
def difference_detail_by_location(request, location_name):
    if not request.user.is_superuser:
        messages.error(request, "您沒有權限訪問此頁面。")
        return redirect('inventory_home')

    sort_by = request.GET.get('sort_by', 'bin')
    order = request.GET.get('order', 'asc')

    queryset = Material.objects.filter(
        location__name=location_name
    ).exclude(
        latest_counted_quantity__isnull=True
    ).exclude(
        system_quantity=F('latest_counted_quantity')
    ).annotate(
        difference=ExpressionWrapper(F('latest_counted_quantity') - F('system_quantity'), output_field=models.IntegerField())
    )

    valid_sort_fields = ['bin', 'material_code', 'material_description', 'system_quantity', 'latest_counted_quantity', 'last_counted_by__username', 'last_counted_date', 'difference']
    if sort_by not in valid_sort_fields:
        sort_by = 'bin'

    order_prefix = '-' if order == 'desc' else ''
    materials = queryset.order_by(f'{order_prefix}{sort_by}')

    context = {
        'location_name': location_name,
        'materials': materials,
        'sort_by': sort_by,
        'order': order,
    }
    return render(request, 'inventory/difference_detail.html', context)

@login_required
def export_differences_excel(request, location_name):
    if not request.user.is_superuser:
        messages.error(request, "您沒有權限執行此操作。")
        return redirect('inventory_home')

    materials = Material.objects.filter(
        location__name=location_name
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

@login_required
def danger_zone(request):
    if not request.user.is_superuser:
        messages.error(request, "您沒有權限訪問此頁面。")
        return redirect('inventory_home')
    return render(request, 'inventory/danger_zone.html')

@login_required
def quick_stocktake_view(request):
    """
    Renders the quick stocktake page.
    """
    if not request.user.is_superuser:
        messages.error(request, "您沒有權限訪問此頁面。")
        return redirect('inventory_home')
    return render(request, 'inventory/quick_stocktake.html')

@login_required
def search_material_for_stocktake(request):
    """
    Searches for a material by its code and returns its details as JSON.
    This is used by the quick stocktake page's AJAX functionality.
    """
    if not request.user.is_superuser:
        return JsonResponse({'status': 'error', 'message': '權限不足'}, status=403)

    material_code = request.GET.get('material_code', '').strip()
    print(f"[DEBUG] Received material_code: {material_code}")  # Logging

    if not material_code:
        print("[DEBUG] material_code is empty.") # Logging
        return JsonResponse({'status': 'error', 'message': '請提供物料號碼'}, status=400)

    try:
        material = Material.objects.get(material_code=material_code)
        print(f"[DEBUG] Found material: {material.material_code}") # Logging
        data = {
            'status': 'success',
            'material': {
                'id': material.id,
                'material_code': material.material_code,
                'material_description': material.material_description,
                'location': material.location.name,
                'bin': material.bin,
                'system_quantity': material.system_quantity,
                'latest_counted_quantity': material.latest_counted_quantity,
                'last_counted_date': material.last_counted_date.isoformat() if material.last_counted_date else None,
                'last_counted_by': material.last_counted_by.username if material.last_counted_by else None,
            }
        }
        return JsonResponse(data)
    except Material.DoesNotExist:
        print(f"[DEBUG] Material with code '{material_code}' not found.") # Logging
        return JsonResponse({'status': 'error', 'message': '找不到該物料號碼'}, status=404)
    except Exception as e:
        print(f"[DEBUG] An exception occurred: {e}") # Logging
        return JsonResponse({'status': 'error', 'message': str(e)}, status=500)


@login_required
def clear_all_materials(request):
    if not request.user.is_superuser:
        messages.error(request, "您沒有權限執行此操作。")
        return redirect('danger_zone')

    if request.method == 'POST':
        confirmation = request.POST.get('confirmation')
        if confirmation == 'DELETE':
            try:
                with transaction.atomic():
                    # Delete all related data first to avoid PROTECT issues
                    StocktakeItem.objects.all().delete()
                    Stocktake.objects.all().delete()
                    # Material deletion will cascade to transactions and images
                    count, _ = Material.objects.all().delete()
                messages.success(request, f"成功清除所有物料資料 (共 {count} 筆)。")
                return redirect('inventory_home')
            except Exception as e:
                messages.error(request, f"清除資料時發生錯誤: {e}")
        else:
            messages.warning(request, "確認文字不符，操作已取消。")
    
    return redirect('danger_zone')