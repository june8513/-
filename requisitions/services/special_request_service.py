from decimal import Decimal
from django.db import transaction
from django.utils import timezone
from django.db.models.functions import Concat
from django.db.models import Q, F, Value
from requisitions.models import (
    UserSelectedMaterial, WorkOrderMaterial, Requisition, 
    RequisitionItem, WorkOrder, ProcessType
)
from django.contrib.auth.models import User

class SpecialRequestService:
    @staticmethod
    def update_user_materials(user, raw_material_text):
        """
        大量更新使用者關注的物料清單
        raw_material_text: 每一行一個物料號碼的字串
        """
        # 解析字串並去除空白與重複，且只取前 10 碼
        material_list = [
            m.strip()[:10] for m in raw_material_text.split('\n') 
            if m.strip()
        ]
        material_list = list(set(material_list))

        with transaction.atomic():
            # 刪除舊有的
            UserSelectedMaterial.objects.filter(user=user).delete()
            
            # 批量建立新的
            objs = [
                UserSelectedMaterial(user=user, material_number=m)
                for m in material_list
            ]
            UserSelectedMaterial.objects.bulk_create(objs)
        
        return len(objs)

    @staticmethod
    def get_matching_materials(user, order_numbers):
        """
        根據工單號碼比對使用者選定的物料
        """
        # 獲取使用者關注的物料清單
        selected_materials = UserSelectedMaterial.objects.filter(user=user).values_list('material_number', flat=True)
        
        if not selected_materials:
            return []

        from django.db.models.functions import Substr, Coalesce, Greatest
        from django.db.models import Sum, OuterRef, Subquery, DecimalField, F
        from decimal import Decimal

        total_confirmed_subquery = RequisitionItem.objects.filter(
            source_material=OuterRef('pk')
        ).values('source_material').annotate(
            total=Sum('confirmed_quantity')
        ).values('total')

        # 搜尋工單中匹配的物料 (比對前 10 碼)，且過濾掉已完全撥料的項目
        matching_items = WorkOrderMaterial.objects.annotate(
            material_number_10=Substr('material_number', 1, 10),
            total_confirmed_quantity=Greatest(
                Coalesce(Subquery(total_confirmed_subquery, output_field=DecimalField()), Decimal('0.0')),
                Coalesce(F('sap_withdrawn_quantity'), Decimal('0.0')),
                Coalesce(F('confirmed_quantity'), Decimal('0.0'))
            )
        ).filter(
            order_number__in=order_numbers,
            material_number_10__in=selected_materials,
            is_active=True
        ).select_related('machine_model', 'process_type')

        return matching_items

    @staticmethod
    def create_special_requisition(user, order_numbers, request_date, demand_person_name, items_data):
        """
        建立特殊的申請單並處理拆分邏輯
        items_data: list of dicts { 'wom_id': int, 'request_qty': decimal }
        """
        # 嘗試尋找對應的使用者 (透過姓名或帳號)
        demand_user = User.objects.annotate(
            full_name=Concat('last_name', 'first_name')
        ).filter(Q(full_name=demand_person_name) | Q(username=demand_person_name)).first()
        
        # 為了簡化，如果找不到精確匹配，就先不填 demand_person (僅存在備註)
        # 或是更進階：尋找 UserProfile 的顯示名稱
        
        with transaction.atomic():
            # 1. 建立申請單 (標題改為申請日期)
            # 為了避免重複，我們加上需求人姓名
            title_date = request_date if isinstance(request_date, str) else request_date.strftime('%Y-%m-%d')
            requisition_title = f"{title_date} - {demand_person_name}"

            process_type_name = "組件"
            # 獲取「組件」投料點物件
            component_process_type = ProcessType.objects.filter(name__icontains=process_type_name).first()

            requisition = Requisition.objects.create(
                order_number=requisition_title,
                applicant=user,
                demand_person=demand_user,
                request_date=request_date,
                process_type=process_type_name,
                requisition_type='finished',
                remarks=f"需求人員: {demand_person_name} (特殊批量申請)"
            )

            for data in items_data:
                wom = WorkOrderMaterial.objects.get(id=data['wom_id'])
                req_qty = Decimal(str(data['request_qty']))

                if req_qty > wom.required_quantity:
                    raise ValueError(f"物料 {wom.material_number} 申請數量 ({req_qty}) 大於工單數量 ({wom.required_quantity})")

                # 如果有找到「組件」投料點，則更新原物料的投料點
                if component_process_type:
                    wom.process_type = component_process_type
                    wom.save()

                final_wom = wom
                # 2. 處理拆分邏輯
                if req_qty < wom.required_quantity:
                    # 拆分：原項減去需求量，新項承接需求量
                    remaining_qty = wom.required_quantity - req_qty
                    wom.required_quantity = remaining_qty
                    wom.save()

                    # 複製一個新的 WOM 來對應這張申請單
                    final_wom = WorkOrderMaterial.objects.create(
                        machine_model=wom.machine_model,
                        order_number=wom.order_number,
                        material_number=wom.material_number,
                        item_name=wom.item_name,
                        required_quantity=req_qty,
                        process_type=component_process_type or wom.process_type,
                        material_type=wom.material_type,
                        estimated_arrival_date=wom.estimated_arrival_date,
                        demand_date=request_date
                    )

                # 3. 建立 RequisitionItem (不進行自動撥料)
                new_item = RequisitionItem.objects.create(
                    requisition=requisition,
                    source_material=final_wom,
                    order_number=final_wom.order_number,
                    material_number=final_wom.material_number,
                    item_name=final_wom.item_name,
                    required_quantity=req_qty,
                    stock_quantity=0,
                    dispatch_status=None
                )

                # 4. 同步更新其他既存的待處理申請單
                # 尋找同一工單、同一物料，且狀態為「需求已提交」或「撥料中」的其他申請單項目
                other_items = RequisitionItem.objects.filter(
                    order_number=wom.order_number,
                    material_number=wom.material_number,
                    requisition__status__in=['demand_submitted', 'dispatch_in_progress'],
                    dispatch_status__isnull=True
                ).exclude(id=new_item.id)

                for other_item in other_items:
                    if other_item.required_quantity > remaining_qty if 'remaining_qty' in locals() else wom.required_quantity:
                        # 如果既存單據需求量大於分拆後的剩餘量，則更新它
                        other_item.required_quantity = remaining_qty if 'remaining_qty' in locals() else wom.required_quantity
                        if other_item.required_quantity <= 0:
                            other_item.delete()
                        else:
                            other_item.save()

            # 5. 申請單狀態保持為「需求已提交」
            requisition.status = 'demand_submitted'
            requisition.save()

            return requisition
