from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.contrib.auth.models import User
from django.contrib import messages
from inventory.models import Material, StorageLocation
from .models import MaterialSpecification
from .forms import MaterialSpecificationForm
import pandas as pd
from django.db import transaction
from decimal import Decimal, InvalidOperation
from django.core.paginator import Paginator
from inventory.forms import MaterialForm

@login_required
def material_spec_list(request):
    query = request.GET.get('q')
    is_admin = request.user.is_superuser

    # Start with a base queryset
    queryset = Material.objects.all().select_related('specification').order_by('material_code')

    if query:
        # If there is a search query, filter based on it (for all users)
        materials_list = queryset.filter(material_code__icontains=query)
        showing_all = False
    elif is_admin:
        # If the user is an admin and there is no query, show all materials
        materials_list = queryset
        showing_all = True
    else:
        # If not an admin and no query, show nothing
        materials_list = Material.objects.none()
        showing_all = False

    paginator = Paginator(materials_list, 20)  # Show 20 materials per page
    page_number = request.GET.get('page')
    materials_page = paginator.get_page(page_number)

    return render(request, 'specifications/material_spec_list.html', {
        'materials': materials_page,
        'is_admin': is_admin,
        'showing_all': showing_all,
        'query': query,
    })

@login_required
def material_spec_edit(request, material_id):
    material = get_object_or_404(Material, pk=material_id)
    spec, created = MaterialSpecification.objects.get_or_create(material=material)

    if request.method == 'POST':
        form = MaterialSpecificationForm(request.POST, request.FILES, instance=spec)
        material_form = MaterialForm(request.POST, instance=material)
        if form.is_valid() and material_form.is_valid():
            form.save()
            material_form.save()
            messages.success(request, '物料規格已成功儲存。')
            return redirect('specifications:material_spec_list')
    else:
        form = MaterialSpecificationForm(instance=spec)
        material_form = MaterialForm(instance=material)

    return render(request, 'specifications/material_spec_edit.html', {'form': form, 'material_form': material_form, 'material': material})

@login_required
def import_material_specs(request):
    if request.method == 'POST' and request.FILES.get('excel_file'):
        excel_file = request.FILES['excel_file']
        try:
            df = pd.read_excel(excel_file)
            
            expected_columns = ['material_code', 'material_description', 'size', 'weight', 'detailed_description']
            if not all(col in df.columns for col in expected_columns):
                missing_cols = ", ".join([col for col in expected_columns if col not in df.columns])
                messages.error(request, f"Excel 檔案缺少必要的欄位，請檢查是否包含: {missing_cols}")
                return redirect('specifications:material_spec_list')

            with transaction.atomic():
                for index, row in df.iterrows():
                    # Clean material_code
                    raw_material_code = row['material_code']
                    material_code = str(raw_material_code)
                    if material_code.endswith('.0'):
                        material_code = material_code[:-2]

                    # Handle the location ForeignKey
                    location_name = row.get('location')
                    location_obj = None
                    if location_name:
                        location_obj, _ = StorageLocation.objects.get_or_create(name=location_name)

                    # --- Data Cleaning ---
                    # Handle system_quantity
                    try:
                        system_quantity_value = int(row.get('system_quantity', 0))
                    except (ValueError, TypeError):
                        system_quantity_value = 0 # Default to 0 if conversion fails

                    # Handle weight
                    try:
                        # Convert to string first to handle potential float values from pandas
                        weight_value = Decimal(str(row.get('weight', '0')))
                    except (InvalidOperation, ValueError, TypeError):
                        weight_value = None # Default to None if conversion fails

                    # --- Database Operations ---
                    material, created = Material.objects.update_or_create(
                        material_code=material_code,
                        defaults={
                            'material_description': row['material_description'],
                            'location': location_obj,
                            'bin': row.get('bin', ''),
                            'system_quantity': system_quantity_value, # Use cleaned value
                        }
                    )

                    MaterialSpecification.objects.update_or_create(
                        material=material,
                        defaults={
                            'size': row['size'],
                            'weight': weight_value, # Use cleaned value
                            'detailed_description': row['detailed_description'],
                        }
                    )
            messages.success(request, f"成功匯入/更新 {len(df)} 筆物料規格。")

        except Exception as e:
            messages.error(request, f"匯入失敗，發生預期外的錯誤: {e}")

    return redirect('specifications:material_spec_list')

@login_required
def redirect_to_material_edit(request):
    if request.method == 'POST':
        material_code = request.POST.get('material_code')
        if material_code:
            try:
                material = Material.objects.get(material_code=material_code)
                return redirect('specifications:material_spec_edit', material_id=material.id)
            except Material.DoesNotExist:
                messages.error(request, f"物料號碼 '{material_code}' 不存在。")
        else:
            messages.error(request, "請輸入物料號碼。")
    return redirect('specifications:material_spec_list')

@login_required
def import_material_purchasers(request):
    if request.method == 'POST' and request.FILES.get('excel_file_purchaser'):
        excel_file = request.FILES['excel_file_purchaser']
        try:
            df = pd.read_excel(excel_file)
            
            expected_columns = ['material_code', 'purchaser']
            if not all(col in df.columns for col in expected_columns):
                missing_cols = ", ".join([col for col in expected_columns if col not in df.columns])
                messages.error(request, f"Excel 檔案缺少必要的欄位，請檢查是否包含: {missing_cols}")
                return redirect('specifications:material_spec_list')

            updated_count = 0
            not_found_materials = []
            not_found_users = []

            with transaction.atomic():
                for index, row in df.iterrows():
                    # Clean material_code
                    raw_material_code = row['material_code']
                    material_code = str(raw_material_code)
                    if material_code.endswith('.0'):
                        material_code = material_code[:-2]
                        
                    purchaser_name = row.get('purchaser', '')

                    try:
                        material = Material.objects.get(material_code=material_code)
                        material.purchaser = purchaser_name
                        material.save()
                        updated_count += 1
                            
                    except Material.DoesNotExist:
                        not_found_materials.append(material_code)

            if updated_count > 0:
                messages.success(request, f"成功更新 {updated_count} 筆物料的採購員。")
            if not_found_materials:
                messages.warning(request, f"找不到以下物料號碼: {', '.join(map(str, set(not_found_materials)))}")


        except Exception as e:
            messages.error(request, f"匯入失敗，發生預期外的錯誤: {e}")

    return redirect('specifications:material_spec_list')

@login_required
def export_material_specs_excel(request):
    """
    匯出物料規格到 Excel
    """
    query = request.GET.get('q')
    is_admin = request.user.is_superuser

    # 開始基礎查詢集
    queryset = Material.objects.all().select_related('specification').order_by('material_code')

    if query:
        # 如果有搜尋查詢，則根據查詢進行過濾（對所有使用者）
        materials_list = queryset.filter(material_code__icontains=query)
    elif is_admin:
        # 如果使用者是管理員且沒有查詢，顯示所有物料
        materials_list = queryset
    else:
        # 如果不是管理員且沒有查詢，不顯示任何內容
        materials_list = Material.objects.none()

    # 準備資料
    data = []
    for material in materials_list:
        data.append({
            '物料號碼': material.material_code,
            '物料說明': material.material_description,
            '採購員': material.purchaser or '',
            '大小': material.specification.size if hasattr(material, 'specification') and material.specification else '',
            '重量': material.specification.weight if hasattr(material, 'specification') and material.specification else '',
            '詳細說明': material.specification.detailed_description if hasattr(material, 'specification') and material.specification else '',
        })

    # 建立 DataFrame
    df = pd.DataFrame(data)

    # 建立 Excel 檔案
    from django.http import HttpResponse
    import io
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='物料規格')
    
    output.seek(0)
    
    # 設定回應
    from django.utils import timezone
    filename = f"material_specifications_{timezone.now().strftime('%Y%m%d_%H%M')}.xlsx"
    response = HttpResponse(
        output.read(),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = f'attachment; filename="{filename}"'
    
    return response