import os
import pandas as pd
from django.db import transaction
from requisitions.models import WorkOrderMaterial, MachineModel, ProcessType
from inventory.models import Material, MaterialTransaction
from decimal import Decimal
from django.conf import settings
import traceback

def process_order_model_excel(excel_file_path):
    """
    Processes an Excel file to upload order and machine model data.
    If an (order_number, machine_model) combination is no longer in the Excel,
    all associated WorkOrderMaterial records are marked as inactive.
    Returns a tuple (created_count, updated_count).
    """
    try:
        try:
            df_upload = pd.read_excel(excel_file_path, dtype=str, engine='openpyxl')
        except Exception as e:
            tb_str = traceback.format_exc()
            raise ValueError(f"讀取 Excel 檔案時發生錯誤: {e}\n{tb_str}")

        df_upload.columns = df_upload.columns.str.strip()

        order_col = '訂單單號' if '訂單單號' in df_upload.columns else '訂單'
        if order_col not in df_upload.columns:
            raise ValueError("上傳的 Excel 檔案中找不到 '訂單單號' 或 '訂單' 欄位。")
        machine_model_col = '機型' if '機型' in df_upload.columns else '物料說明'
        if machine_model_col not in df_upload.columns:
            raise ValueError("上傳的 Excel 檔案中找不到 '機型' 或 '物料說明' 欄位。")

        created_count = 0
        updated_count = 0

        # Collect all unique (order_number, machine_model_name) from the new Excel
        new_excel_combinations = set()
        for _, row in df_upload.iterrows():
            order_number = str(row.get(order_col)).strip()
            machine_model_name = str(row.get(machine_model_col, '')).strip()
            if order_number and machine_model_name:
                new_excel_combinations.add((order_number, machine_model_name))

        with transaction.atomic():
            # Deactivate existing WorkOrderMaterial records that are not in the new Excel
            # Iterate through all currently active combinations in the DB
            existing_active_combinations_in_db = set(WorkOrderMaterial.objects.filter(is_active=True).values_list(
                'order_number', 'machine_model__name'
            )) # FIX: Convert to set directly

            for db_order_num, db_machine_model_name in existing_active_combinations_in_db:
                if (db_order_num, db_machine_model_name) not in new_excel_combinations:
                    # This combination is in DB but not in new Excel, so deactivate all its materials
                    WorkOrderMaterial.objects.filter(
                        order_number=db_order_num,
                        machine_model__name=db_machine_model_name
                    ).update(is_active=False)

            # Process the new Excel data (update_or_create)
            for _, row in df_upload.iterrows():
                order_number = str(row.get(order_col)).strip()
                machine_model_name = str(row.get(machine_model_col, '')).strip()

                if not all([order_number, machine_model_name]):
                    continue

                machine_model_obj, _ = MachineModel.objects.get_or_create(name=machine_model_name)

                parent_scope_material_number = "PARENT_SCOPE" 

                defaults = {
                    'item_name': '訂單機型範圍',
                    'required_quantity': Decimal('0.00'),
                    'process_type': None,
                    'is_active': True, # Ensure new/updated materials are active
                }

                existing_parent_scope = WorkOrderMaterial.objects.filter(
                    order_number=order_number,
                    material_number=parent_scope_material_number
                ).first()

                if existing_parent_scope and existing_parent_scope.machine_model != machine_model_obj:
                    raise ValueError(f"訂單 {order_number} 已存在不同的機型 ({existing_parent_scope.machine_model.name})。一個訂單只能有一個機型。")

                obj, created = WorkOrderMaterial.objects.update_or_create(
                    order_number=order_number,
                    material_number=parent_scope_material_number,
                    machine_model=machine_model_obj,
                    defaults=defaults
                )

                if created:
                    created_count += 1
                else:
                    updated_count += 1
        
        return created_count, updated_count

    except Exception as e:
        tb_str = traceback.format_exc()
        raise type(e)(f"處理 Excel 內容時發生未預期的錯誤: {e}\n{tb_str}")

def process_material_details_excel(excel_file_path, required_qty_col):
    """
    Processes an Excel file to upload material details.
    Returns a tuple (created_count, updated_count, deleted_count).
    """
    try:
        # Step 1: Read the process type mapping from the local DB file
        try:
            db_path = os.path.join(settings.BASE_DIR, 'output.xlsx')
            excel_sheets = pd.read_excel(db_path, engine='openpyxl', sheet_name=None)
            df_db = pd.concat(excel_sheets.values(), ignore_index=True)

            if '物料' not in df_db.columns or '機型' not in df_db.columns or '投料點' not in df_db.columns:
                raise ValueError("output.xlsx 檔案中必須包含 '物料', '機型','投料點' 欄位。")

            df_db['material_prefix'] = df_db['物料'].astype(str).str[:10]
            df_db['machine_model_name'] = df_db['機型'].astype(str).str.strip()

            df_db['composite_key'] = list(zip(df_db['material_prefix'], df_db['machine_model_name']))

            process_type_map = df_db.set_index('composite_key')['投料點'].to_dict()

        except Exception as e:
            raise ValueError(f"讀取投料點資料庫 (output.xlsx) 時發生錯誤:{e}")

        # Step 2: Read the uploaded Excel file
        df_upload = pd.read_excel(excel_file_path, dtype=str, engine='openpyxl')
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

        updated_count = 0
        created_count = 0
        deleted_count = 0

        with transaction.atomic():
            materials_to_create = []
            materials_to_update = []
            uploaded_material_keys = set()
            created_material_keys = set()

            all_order_numbers_in_upload = df_upload[order_col].astype(str).str.strip().unique()
            existing_materials_db = WorkOrderMaterial.objects.filter(order_number__in=all_order_numbers_in_upload).exclude(material_number="PARENT_SCOPE").select_related('machine_model')

            existing_materials_lookup = {}
            for material in existing_materials_db:
                key = (str(material.order_number).strip(), str(material.material_number).strip(), str(material.machine_model.name).strip())
                existing_materials_lookup[key] = material

            for _, row in df_upload.iterrows():
                order_number_clean = str(row.get(order_col)).strip()
                material_number_clean = str(row.get('物料')).strip()

                if not all([order_number_clean, material_number_clean]):
                    continue

                parent_scope_entry = WorkOrderMaterial.objects.filter(
                    order_number=order_number_clean,
                    material_number="PARENT_SCOPE"
                ).first()

                if not parent_scope_entry or not parent_scope_entry.machine_model:
                    raise ValueError(f"訂單 {order_number_clean}的父階範圍不存在或缺少機型資訊。請先上傳訂單與機型 Excel。")

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

                current_material_key = (order_number_clean, material_number_clean, machine_model_name_clean)
                uploaded_material_keys.add(current_material_key)

                if current_material_key in existing_materials_lookup:
                    material_instance = existing_materials_lookup[current_material_key]
                    material_instance.item_name = item_name_clean
                    material_instance.required_quantity = required_quantity_clean
                    material_instance.process_type = process_type_obj
                    materials_to_update.append(material_instance)
                    updated_count += 1
                else:
                    if current_material_key not in created_material_keys:
                        materials_to_create.append(
                            WorkOrderMaterial(
                                order_number=order_number_clean,
                                material_number=material_number_clean,
                                machine_model=machine_model_obj,
                                item_name=item_name_clean,
                                required_quantity=required_quantity_clean,
                                process_type=process_type_obj
                            )
                        )
                        created_material_keys.add(current_material_key)
                        created_count += 1

            if materials_to_create:
                WorkOrderMaterial.objects.bulk_create(materials_to_create)
            if materials_to_update:
                WorkOrderMaterial.objects.bulk_update(materials_to_update, fields=['item_name', 'required_quantity', 'process_type'])

            uploaded_deletion_scopes = set()
            for _, row in df_upload.iterrows():
                order_number = str(row.get(order_col)).strip()

                parent_scope_entry_for_deletion = WorkOrderMaterial.objects.filter(
                    order_number=order_number,
                    material_number="PARENT_SCOPE"
                ).first()

                if not parent_scope_entry_for_deletion or not parent_scope_entry_for_deletion.machine_model:
                    continue

                machine_model_name_for_deletion = parent_scope_entry_for_deletion.machine_model.name
                uploaded_deletion_scopes.add((order_number, machine_model_name_for_deletion))

            for order_num, model_name in uploaded_deletion_scopes:
                existing_materials_in_scope = WorkOrderMaterial.objects.filter(
                    order_number=order_num,
                    machine_model__name=model_name
                ).exclude(material_number="PARENT_SCOPE")

                for material in existing_materials_in_scope:
                    db_key = (str(material.order_number).strip(), str(material.material_number).strip(), str(material.machine_model.name).strip())
                    if db_key not in uploaded_material_keys:
                        material.delete()
                        deleted_count += 1
        
        return created_count, updated_count, deleted_count

    except Exception as e:
        raise e

def process_inventory_excel(excel_file_path):
    """
    Processes an Excel file to upload inventory data, ignoring location and bin.
    Returns a tuple (created_count, updated_count).
    """
    try:
        expected_columns = ['物料', '物料說明', '未限制']
        try:
            df = pd.read_excel(excel_file_path, dtype={'物料': str})
        except Exception as e:
            tb_str = traceback.format_exc()
            raise ValueError(f"讀取庫存 Excel 檔案時發生錯誤: {e}\n{tb_str}")

        if not all(col in df.columns for col in expected_columns):
            missing_cols = ", ".join([col for col in expected_columns if col not in df.columns])
            raise ValueError(f"庫存 Excel 檔案缺少必要的欄位: {missing_cols}")

        created_count = 0
        updated_count = 0

        with transaction.atomic():
            for _, row in df.iterrows():
                material_code = row.get('物料')
                if not material_code or pd.isna(material_code):
                    continue

                material_code = str(material_code).strip()

                defaults = {
                    'material_description': row.get('物料說明', ''),
                    'system_quantity': pd.to_numeric(row.get('未限制'), errors='coerce') or 0,
                }

                material, created = Material.objects.update_or_create(
                    material_code=material_code,
                    defaults=defaults
                )
                
                # Note: Creating a MaterialTransaction is omitted here because there is no
                # 'user' in an automated context.

                if created:
                    created_count += 1
                else:
                    updated_count += 1
        
        return created_count, updated_count

    except Exception as e:
        tb_str = traceback.format_exc()
        raise type(e)(f"處理庫存 Excel 內容時發生未預期的錯誤: {e}\n{tb_str}")