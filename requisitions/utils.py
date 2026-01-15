import os
import pandas as pd
import requests # Added requests
from django.db import transaction
from django.db.models import Q # Added Q
from requisitions.models import WorkOrder, WorkOrderMaterial, MachineModel, ProcessType, Requisition
from inventory.models import Material, MaterialTransaction
from decimal import Decimal
from django.conf import settings
from django.utils import timezone # Added timezone
import traceback

def notify_requisition_shortages(requisition):
    """
    Sends a consolidated notification about ALL shortages for a given requisition.
    Triggered after batch allocation updates.
    """
    try:
        # Get all items for this requisition that have shortages
        # We fetch all items to calculate the full picture if needed, or just filter shortages
        # Let's send a list of items that are currently short
        items = requisition.items.all()
        
        shortage_list = []
        has_shortage = False
        
        for item in items:
            confirmed = item.confirmed_quantity or Decimal('0')
            shortage = item.required_quantity - confirmed
            
            if shortage > 0:
                has_shortage = True
                shortage_list.append({
                    "material_number": item.material_number,
                    "item_name": item.item_name,
                    "required_quantity": float(item.required_quantity),
                    "confirmed_quantity": float(confirmed),
                    "shortage_quantity": float(shortage),
                    "status": "OPEN"
                })

        # Overall status for the requisition
        overall_status = "OPEN" if has_shortage else "RESOLVED"

        payload = {
            "order_number": requisition.order_number,
            "process_type": requisition.process_type,
            "requisition_id": requisition.pk,
            "status": overall_status,
            "timestamp": timezone.now().isoformat(),
            "shortage_items": shortage_list
        }
        
        # Send Request
        external_url = getattr(settings, 'SHORTAGE_NOTIFICATION_URL', None)
        
        if external_url:
            try:
                response = requests.post(external_url, json=payload, timeout=5)
                response.raise_for_status()
                print(f"[Requisition Notification] Success! Sent {len(shortage_list)} items for Order {requisition.order_number}")
            except requests.RequestException as req_err:
                print(f"[Requisition Notification] Request failed: {req_err}")
        else:
            print(f"[Requisition Notification] URL not configured. Payload: {payload}")
            
    except Exception as e:
        print(f"[Requisition Notification] Error: {e}")

def process_order_model_excel(excel_file_path):
    """
    Processes an Excel file to upload order and machine model data.
    If an (order_number, machine_model) combination is no longer in the Excel,
    it sets a status message on the corresponding WorkOrder.
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
            # Get all unique order numbers from the upload
            all_order_numbers_in_upload = {str(row.get(order_col)).strip() for _, row in df_upload.iterrows() if str(row.get(order_col)).strip()}

            # Create WorkOrder entries for all unique order numbers in the upload, and update shipping date
            for order_number in all_order_numbers_in_upload:
                # Find the first row for this order number to get the shipping date
                order_rows = df_upload[df_upload[order_col] == order_number]
                shipping_date = None
                if not order_rows.empty and '出貨日期' in df_upload.columns:
                    date_val = order_rows.iloc[0]['出貨日期']
                    if pd.notna(date_val):
                        try:
                            shipping_date = pd.to_datetime(date_val).date()
                        except (ValueError, TypeError):
                            shipping_date = None # Ignore if parsing fails
                
                WorkOrder.objects.update_or_create(
                    order_number=order_number,
                    defaults={'shipping_date': shipping_date}
                )

            # Identify combinations in DB that are not in the new Excel
            existing_active_combinations_in_db = set(WorkOrderMaterial.objects.filter(is_active=True).values_list(
                'order_number', 'machine_model__name'
            ))

            deactivated_order_numbers = set()
            for db_order_num, db_machine_model_name in existing_active_combinations_in_db:
                if (db_order_num, db_machine_model_name) not in new_excel_combinations:
                    # This combination is in DB but not in new Excel.
                    # Instead of deactivating, we set a message on the WorkOrder.
                    deactivated_order_numbers.add(db_order_num)

            # Set status message for WorkOrders that had changes
            if deactivated_order_numbers:
                for order_num in deactivated_order_numbers:
                    WorkOrder.objects.filter(order_number=order_num).update(
                        status_message="訂單/機型組合已變更，請確認"
                    )

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

def _update_requisition_alert(order_number, process_type_name, message, is_demand_increase=False):
    """
    Helper to update requisition alert status and potentially revert status if demand increases.
    """
    try:
        # Filter for the exact requisition
        # Note: There should be only one per order/process_type due to UniqueConstraint
        reqs = Requisition.objects.filter(order_number=order_number, process_type=process_type_name)
        
        for req in reqs:
            # Append message
            timestamp = timezone.now().strftime('%Y-%m-%d %H:%M')
            new_msg = f"[{timestamp}] {message}"
            if req.alert_message:
                req.alert_message += f"\n{new_msg}"
            else:
                req.alert_message = new_msg
            
            req.has_alert = True
            
            # Revert status if demand increases (and it was completed/signed_off/archived)
            if is_demand_increase:
                if req.status in ['dispatch_completed', 'signed_off', 'archived'] or req.is_archived:
                    req.status = 'dispatch_in_progress'
                    req.is_archived = False # Un-archive
                    # We might want to keep dispatch_performed as True if partial dispatch is done?
                    # But strictly, if demand increased, dispatch is NOT fully performed.
                    req.dispatch_performed = False 
            
            req.save()
    except Exception as e:
        print(f"Error updating alert for {order_number}: {e}")

def process_material_details_excel(excel_file_path, required_qty_col):
    """
    Processes an Excel file to upload material details.
    Returns a tuple (created_count, updated_count, deactivated_count).
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

        # Determine parent description column
        parent_desc_col = '上層物料說明' if '上層物料說明' in df_upload.columns else None

        # --- START of FIX: Aggregate data before processing ---
        group_cols = [order_col, '物料']
        if parent_desc_col:
            group_cols.append(parent_desc_col)

        df_aggregated = df_upload.groupby(group_cols).agg({
            required_qty_col: 'sum',
            '物料說明': 'first'  # Keep the first item name found
        }).reset_index()
        # --- END of FIX ---

        updated_count = 0
        created_count = 0
        deactivated_count = 0

        with transaction.atomic():
            materials_to_create = []
            materials_to_update = []
            uploaded_material_keys = set()
            created_material_keys = set()

            all_order_numbers_in_upload = df_upload[order_col].astype(str).str.strip().unique()
            
            # Ensure WorkOrder entries exist for all order numbers in the upload
            for order_number in all_order_numbers_in_upload:
                WorkOrder.objects.get_or_create(order_number=order_number)

            # --- OPTIMIZATION: Pre-fetch Parent Scopes ---
            parent_scope_map = {}
            parent_scopes = WorkOrderMaterial.objects.filter(
                order_number__in=all_order_numbers_in_upload,
                material_number="PARENT_SCOPE"
            ).select_related('machine_model')
            
            for ps in parent_scopes:
                if ps.machine_model:
                    parent_scope_map[ps.order_number] = ps.machine_model

            # --- OPTIMIZATION: Pre-fetch/Create Process Types ---
            needed_process_types = set()
            
            # First pass to identify needed process types
            for _, row in df_aggregated.iterrows():
                order_number_clean = str(row.get(order_col)).strip()
                material_number_clean = str(row.get('物料')).strip()
                
                if not all([order_number_clean, material_number_clean]):
                    continue
                    
                machine_model_obj = parent_scope_map.get(order_number_clean)
                if not machine_model_obj:
                    continue # Will be caught as error in main loop
                    
                material_prefix = material_number_clean[:10]
                composite_lookup_key = (material_prefix, machine_model_obj.name)
                process_type_name = str(process_type_map.get(composite_lookup_key, '其他')).strip()
                
                needed_process_types.add((process_type_name, machine_model_obj))
            
            process_type_cache = {} # (name, machine_model_id) -> obj
            
            if needed_process_types:
                q_objects = Q()
                for pt_name, mm_obj in needed_process_types:
                    q_objects |= Q(name=pt_name, machine_model=mm_obj)
                
                existing_pts = ProcessType.objects.filter(q_objects)
                for pt in existing_pts:
                    process_type_cache[(pt.name, pt.machine_model_id)] = pt
                
                pts_to_create = []
                for pt_name, mm_obj in needed_process_types:
                    if (pt_name, mm_obj.id) not in process_type_cache:
                        pts_to_create.append(ProcessType(name=pt_name, machine_model=mm_obj))
                
                if pts_to_create:
                    # Ignore conflicts just in case of race condition, though atomic block helps
                    created_pts = ProcessType.objects.bulk_create(pts_to_create, ignore_conflicts=True)
                    # Re-fetch or manually add to cache. Bulk create with ignore_conflicts might not return objs on some DBs
                    # Safer to re-fetch all needed again
                    existing_pts_refresh = ProcessType.objects.filter(q_objects)
                    for pt in existing_pts_refresh:
                        process_type_cache[(pt.name, pt.machine_model_id)] = pt

            existing_materials_db = WorkOrderMaterial.objects.filter(order_number__in=all_order_numbers_in_upload).exclude(material_number="PARENT_SCOPE").select_related('machine_model')

            existing_materials_lookup = {}
            for material in existing_materials_db:
                key = (str(material.order_number).strip(), str(material.material_number).strip(), str(material.machine_model.name).strip())
                existing_materials_lookup[key] = material

            for _, row in df_aggregated.iterrows():
                order_number_clean = str(row.get(order_col)).strip()
                material_number_clean = str(row.get('物料')).strip()

                if not all([order_number_clean, material_number_clean]):
                    continue

                machine_model_obj = parent_scope_map.get(order_number_clean)

                if not machine_model_obj:
                    raise ValueError(f"訂單 {order_number_clean}的父階範圍不存在或缺少機型資訊。請先上傳訂單與機型 Excel。")

                machine_model_name_clean = machine_model_obj.name

                material_prefix = material_number_clean[:10]
                parent_desc_val = str(row.get(parent_desc_col, '')).strip() if parent_desc_col else ""
                
                # Check for learned rule
                from requisitions.models import MaterialProcessTypeRule
                rules = MaterialProcessTypeRule.objects.filter(
                    material_prefix=material_prefix,
                    machine_model_name=machine_model_name_clean
                ).order_by('-parent_material_desc_keyword') # Empty strings last (or first depending on DB, but non-empty usually longer)
                # Actually, '' is shorter than 'keyword', so '-' length might be better, or just rely on logic
                # Let's filter in python
                
                process_type_name = None
                
                # Try to find specific keyword match first
                for rule in rules:
                    if rule.parent_material_desc_keyword and rule.parent_material_desc_keyword in parent_desc_val:
                        process_type_name = rule.process_type_name
                        break
                
                # If no keyword match, look for generic rule (empty keyword)
                if not process_type_name:
                    for rule in rules:
                        if not rule.parent_material_desc_keyword:
                            process_type_name = rule.process_type_name
                            break
                            
                if not process_type_name:
                    composite_lookup_key = (material_prefix, machine_model_name_clean)
                    process_type_name = str(process_type_map.get(composite_lookup_key, '其他')).strip()

                # Use cache
                process_type_obj = process_type_cache.get((process_type_name, machine_model_obj.id))
                
                if not process_type_obj:
                     # Fallback if somehow missed in pre-fetch (shouldn't happen)
                     process_type_obj, _ = ProcessType.objects.get_or_create(
                        name=process_type_name,
                        machine_model=machine_model_obj
                    )

                raw_item_name = row.get('物料說明', '')
                if pd.isna(raw_item_name) or str(raw_item_name).lower() == 'nan':
                    item_name_clean = ""
                else:
                    item_name_clean = str(raw_item_name).strip()

                required_quantity_clean = row.get(required_qty_col, 0)

                current_material_key = (order_number_clean, material_number_clean, machine_model_name_clean)
                uploaded_material_keys.add(current_material_key)

                if current_material_key in existing_materials_lookup:
                    material_instance = existing_materials_lookup[current_material_key]
                    
                    is_dirty = False
                    
                    # String comparison with handling for None
                    db_name = str(material_instance.item_name or "").strip()
                    excel_name = str(item_name_clean).strip()
                    if db_name != excel_name:
                        material_instance.item_name = excel_name
                        is_dirty = True
                    
                    # Quantity comparison
                    db_qty = 0.0
                    excel_qty = 0.0
                    try:
                        db_qty = float(material_instance.required_quantity or 0)
                        excel_qty = float(required_quantity_clean or 0)
                        if abs(db_qty - excel_qty) > 0.0001:
                            material_instance.required_quantity = required_quantity_clean
                            is_dirty = True
                            
                            diff = excel_qty - db_qty
                            msg_type = "增加" if diff > 0 else "減少"
                            _update_requisition_alert(
                                order_number_clean, 
                                process_type_obj.name, 
                                f"物料 {material_number_clean} 需求{msg_type}: {db_qty:.2f} -> {excel_qty:.2f}",
                                is_demand_increase=(diff > 0)
                            )
                    except (ValueError, TypeError):
                        if str(material_instance.required_quantity) != str(required_quantity_clean):
                             material_instance.required_quantity = required_quantity_clean
                             is_dirty = True
                    
                    # Process Type comparison (ID)
                    if material_instance.process_type_id != process_type_obj.id:
                        material_instance.process_type = process_type_obj
                        is_dirty = True
                    
                    # Active status
                    if not material_instance.is_active:
                        material_instance.is_active = True
                        is_dirty = True

                    if is_dirty:
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
                                process_type=process_type_obj,
                                is_active=True
                            )
                        )
                        created_material_keys.add(current_material_key)
                        created_count += 1
                        
                        _update_requisition_alert(
                            order_number_clean, 
                            process_type_obj.name, 
                            f"新增物料: {material_number_clean} (需求: {required_quantity_clean})",
                            is_demand_increase=True
                        )

            if materials_to_create:
                WorkOrderMaterial.objects.bulk_create(materials_to_create, batch_size=2000)
            if materials_to_update:
                WorkOrderMaterial.objects.bulk_update(materials_to_update, fields=['item_name', 'required_quantity', 'process_type', 'is_active'], batch_size=2000)

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
                ).exclude(material_number="PARENT_SCOPE").select_related('process_type')

                materials_to_deactivate = []
                for material in existing_materials_in_scope:
                    db_key = (str(material.order_number).strip(), str(material.material_number).strip(), str(material.machine_model.name).strip())
                    if db_key not in uploaded_material_keys:
                        material.is_active = False
                        materials_to_deactivate.append(material)
                        deactivated_count += 1
                        
                        pt_name = material.process_type.name if material.process_type else "Unknown"
                        _update_requisition_alert(
                            order_num, 
                            pt_name, 
                            f"移除物料: {material.material_number} (原需求: {material.required_quantity})",
                            is_demand_increase=False
                        )
                        
                        # Sync RequisitionItems
                        from requisitions.models import RequisitionItem
                        req_items = RequisitionItem.objects.filter(source_material=material)

                        for item in req_items:
                            confirmed = item.confirmed_quantity or Decimal('0')
                            if confirmed > 0:
                                # Has dispatch: Close the shortage by setting required = confirmed
                                item.required_quantity = confirmed
                                if "(已刪除)" not in item.item_name:
                                    item.item_name += " (已刪除)"
                                item.dispatch_status = 'dispatched'
                                item.save()
                            else:
                                # No dispatch: Delete item completely
                                item.delete()
                
                if materials_to_deactivate:
                    WorkOrderMaterial.objects.bulk_update(materials_to_deactivate, ['is_active'])
                    # Set status message for the corresponding WorkOrder
                    WorkOrder.objects.filter(order_number=order_num).update(
                        status_message="部分物料已不存在於最新匯入資料中，請確認"
                    )
        
        return created_count, updated_count, deactivated_count

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