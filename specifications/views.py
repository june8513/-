from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from inventory.models import Material
from .models import MaterialSpecification
from .forms import MaterialSpecificationForm
import pandas as pd
from django.db import transaction

@login_required
def material_spec_list(request):
    query = request.GET.get('q')
    materials = Material.objects.none()  # Return no materials by default
    if query:
        # Using select_related to fetch the specification along with the material
        # to avoid extra database queries in the template.
        materials = Material.objects.filter(material_code__icontains=query).select_related('specification')
    return render(request, 'specifications/material_spec_list.html', {'materials': materials})

@login_required
def material_spec_edit(request, material_id):
    material = get_object_or_404(Material, pk=material_id)
    spec, created = MaterialSpecification.objects.get_or_create(material=material)

    if request.method == 'POST':
        form = MaterialSpecificationForm(request.POST, request.FILES, instance=spec)
        if form.is_valid():
            form.save()
            messages.success(request, '物料規格已成功儲存。')
            return redirect('material_spec_list')
    else:
        form = MaterialSpecificationForm(instance=spec)

    return render(request, 'specifications/material_spec_edit.html', {'form': form, 'material': material})

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
                return redirect('material_spec_list')

            with transaction.atomic():
                for index, row in df.iterrows():
                    material_code = row['material_code']
                    
                    material, created = Material.objects.update_or_create(
                        material_code=material_code,
                        defaults={
                            'material_description': row['material_description'],
                            'location': row.get('location', ''),
                            'bin': row.get('bin', ''),
                            'system_quantity': row.get('system_quantity', 0),
                        }
                    )
                    
                    MaterialSpecification.objects.update_or_create(
                        material=material,
                        defaults={
                            'size': row['size'],
                            'weight': row['weight'],
                            'detailed_description': row['detailed_description'],
                        }
                    )
            messages.success(request, f"成功匯入/更新 {len(df)} 筆物料規格。")

        except Exception as e:
            messages.error(request, f"匯入失敗，發生預期外的錯誤: {e}")

    return redirect('material_spec_list')

@login_required
def redirect_to_material_edit(request):
    if request.method == 'POST':
        material_code = request.POST.get('material_code')
        if material_code:
            try:
                material = Material.objects.get(material_code=material_code)
                return redirect('material_spec_edit', material_id=material.id)
            except Material.DoesNotExist:
                messages.error(request, f"物料號碼 '{material_code}' 不存在。")
        else:
            messages.error(request, "請輸入物料號碼。")
    return redirect('material_spec_list')