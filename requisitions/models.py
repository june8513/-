from django.db import models
from django.contrib.auth.models import User
from django.db.models import UniqueConstraint

class MachineModel(models.Model):
    name = models.CharField(max_length=100, unique=True, verbose_name="機型名稱")

    class Meta:
        verbose_name = "機型"
        verbose_name_plural = "機型"
        ordering = ['name']

    def __str__(self):
        return self.name

class ProcessType(models.Model):
    name = models.CharField(max_length=100, verbose_name="投料點名稱")
    machine_model = models.ForeignKey(MachineModel, on_delete=models.CASCADE, related_name="process_types", verbose_name="所屬機型")
    parent = models.ForeignKey('self', on_delete=models.SET_NULL, null=True, blank=True, related_name='children', verbose_name="上層投料點")
    is_kit = models.BooleanField(default=False, verbose_name="是否為台份(Kit)")

    class Meta:
        verbose_name = "投料點"
        verbose_name_plural = "投料點"
        unique_together = ('name', 'machine_model')
        ordering = ['machine_model', 'name']

    def __str__(self):
        return self.name

class WorkOrder(models.Model):
    order_number = models.CharField(max_length=100, unique=True, verbose_name="工單單號")
    shipping_date = models.DateField(null=True, blank=True, verbose_name="出貨日期")
    is_archived = models.BooleanField(default=False, verbose_name="是否已歸檔")
    status_message = models.CharField(max_length=255, blank=True, null=True, verbose_name="工單現況")
    created_at = models.DateTimeField(auto_now_add=True, verbose_name="建立時間")
    updated_at = models.DateTimeField(auto_now=True, verbose_name="更新時間")

    class Meta:
        verbose_name = "工單"
        verbose_name_plural = "工單"
        ordering = ['-updated_at']

    def __str__(self):
        return self.order_number

class Requisition(models.Model):
    STATUS_CHOICES = [
        ('demand_submitted', '需求已提交'),
        ('dispatch_in_progress', '撥料中'),
        ('dispatch_completed', '撥料已完成'),
        ('signed_off', '已簽收'),
        ('archived', '已歸檔'),
    ]
    
    REQUISITION_TYPE_CHOICES = [
        ('finished', '成品'),
        ('semi_finished', '半成品'),
    ]

    order_number = models.CharField(max_length=100, verbose_name="訂單單號")
    applicant = models.ForeignKey(User, on_delete=models.CASCADE, related_name='requisitions_applied', verbose_name="申請人")
    request_date = models.DateField(verbose_name="需求日期", db_index=True)
    process_type = models.CharField(max_length=100, verbose_name="需求流程", db_index=True, null=True, blank=True)
    requisition_type = models.CharField(max_length=20, choices=REQUISITION_TYPE_CHOICES, default='finished', verbose_name="申請單類型", db_index=True)
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='demand_submitted', verbose_name="狀態", db_index=True)
    dispatch_performed = models.BooleanField(default=False, verbose_name="已執行撥料")
    is_archived = models.BooleanField(default=False, verbose_name="是否已歸檔")
    
    # Demand Change Alerts
    has_alert = models.BooleanField(default=False, verbose_name="有需求變更警示")
    alert_message = models.TextField(blank=True, null=True, verbose_name="警示訊息")
    
    material_confirmed_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='requisitions_material_confirmed', verbose_name="物料確認人員")
    material_confirmed_date = models.DateTimeField(null=True, blank=True, verbose_name="物料確認日期")
    sign_off_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='requisitions_signed_off', verbose_name="最終簽收人員")
    sign_off_date = models.DateTimeField(null=True, blank=True, verbose_name="最終簽收日期")


    remarks = models.TextField(blank=True, null=True, verbose_name="備註")

    created_at = models.DateTimeField(auto_now_add=True, verbose_name="建立時間")
    updated_at = models.DateTimeField(auto_now=True, verbose_name="更新時間")

    class Meta:
        verbose_name = "撥料申請單"
        verbose_name_plural = "撥料申請單"
        ordering = ['-created_at']
        constraints = [
            UniqueConstraint(fields=['order_number', 'process_type', 'requisition_type'], name='unique_order_per_process_type')
        ]

    def __str__(self):
        return f"撥料申請單: {self.order_number} ({self.process_type}) - {self.applicant.username}"



class WorkOrderMaterial(models.Model):
    MATERIAL_TYPE_CHOICES = [
        ('finished', '成品'),
        ('semi_finished', '半成品'),
    ]
    
    machine_model = models.ForeignKey(MachineModel, on_delete=models.CASCADE, verbose_name="機型", related_name="work_order_materials", null=True, blank=True)
    order_number = models.CharField(max_length=100, db_index=True, verbose_name="訂單單號")
    material_number = models.CharField(max_length=100, db_index=True, verbose_name="物料", null=True, blank=True)
    item_name = models.CharField(max_length=255, verbose_name="品名", null=True, blank=True)
    required_quantity = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="需求數量")
    process_type = models.ForeignKey(ProcessType, on_delete=models.SET_NULL, null=True, blank=True, verbose_name="投料點")
    
    # 新增：區分成品/半成品
    material_type = models.CharField(max_length=20, choices=MATERIAL_TYPE_CHOICES, default='finished', verbose_name="物料類型", db_index=True)
    
    confirmed_quantity = models.DecimalField(max_digits=10, decimal_places=2, null=True, blank=True, verbose_name="已撥料數量")
    is_signed_off = models.BooleanField(default=False, verbose_name="已簽收")
    is_active = models.BooleanField(default=True, verbose_name="是否啟用") # New field
    estimated_arrival_date = models.DateField(null=True, blank=True, verbose_name="預計入料日期")
    demand_date = models.DateField(null=True, blank=True, verbose_name="需求日期") # New field

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "訂單主物料清單"
        verbose_name_plural = "訂單主物料清單"
        ordering = ['order_number', 'process_type', 'material_number']
        

    def __str__(self):
        return f"{self.order_number} - {self.machine_model} - {self.material_number} ({self.item_name})"

class RequisitionItem(models.Model):
    DISPATCH_STATUS_CHOICES = [
        ('dispatched', '已撥料'),
        ('backordered', '欠料'),
    ]

    requisition = models.ForeignKey(Requisition, on_delete=models.CASCADE, related_name='items', verbose_name="所屬撥料單", null=True, blank=True)
    source_material = models.ForeignKey(WorkOrderMaterial, on_delete=models.SET_NULL, null=True, blank=True, related_name='requisition_items', verbose_name="來源主物料")
    order_number = models.CharField(max_length=100, verbose_name="訂單單號")
    material_number = models.CharField(max_length=100, verbose_name="物料", db_index=True)
    item_name = models.CharField(max_length=255, verbose_name="品名")
    required_quantity = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="需求數量")
    stock_quantity = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="庫存數量")
    storage_bin = models.CharField(max_length=100, blank=True, null=True, verbose_name="儲格")
    
    confirmed_quantity = models.DecimalField(max_digits=10, decimal_places=2, null=True, blank=True, verbose_name="確認撥料數量")
    is_signed_off = models.BooleanField(default=False, verbose_name="最終簽收已確認")
    sign_off_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='requisition_items_signed_off', verbose_name="簽收人員")
    sign_off_date = models.DateTimeField(null=True, blank=True, verbose_name="簽收日期")
    dispatch_status = models.CharField(max_length=20, choices=DISPATCH_STATUS_CHOICES, null=True, blank=True, verbose_name="撥料狀態")
    
    # 補撥相關欄位
    is_supplementary = models.BooleanField(default=False, verbose_name="是否為補撥")
    parent_item = models.ForeignKey('self', on_delete=models.SET_NULL, null=True, blank=True, related_name='supplementary_items', verbose_name="原始項目")

    class Meta:
        verbose_name = "撥料物料明細"
        verbose_name_plural = "撥料物料明細"

    def __str__(self):
        return f"{self.item_name} ({self.required_quantity})"

class Inventory(models.Model):
    material_number = models.CharField(max_length=100, unique=True, verbose_name="物料")
    storage_bin = models.CharField(max_length=100, blank=True, null=True, verbose_name="儲格")
    stock_quantity = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="庫存數量")

    class Meta:
        verbose_name = "庫存"
        verbose_name_plural = "庫存"

    def __str__(self):
        return f"{self.material_number} - {self.storage_bin} ({self.stock_quantity})"

class RequisitionImage(models.Model):
    requisition = models.ForeignKey(Requisition, on_delete=models.CASCADE, related_name='images', verbose_name="所屬撥料單")
    image = models.ImageField(upload_to='requisition_images/', verbose_name="撥料單圖片")
    uploaded_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, verbose_name="上傳人員")
    uploaded_at = models.DateTimeField(auto_now_add=True, verbose_name="上傳時間")

    class Meta:
        verbose_name = "撥料單圖片"
        verbose_name_plural = "撥料單圖片"
        ordering = ['-uploaded_at']

    def __str__(self):
        return f"圖片 for {self.requisition.order_number} ({self.uploaded_at.strftime('%Y-%m-%d %H:%M')})"

class WorkOrderMaterialTransaction(models.Model):
    TRANSACTION_TYPES = [
        ('ALLOCATION', '撥料'),
        ('RETURN', '退料'),
        ('MANUAL_UPDATE', '手動修改'),
    ]

    work_order_material = models.ForeignKey(WorkOrderMaterial, on_delete=models.CASCADE, related_name='transactions', verbose_name="訂單物料")
    user = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, verbose_name="操作人員")
    transaction_type = models.CharField(max_length=20, choices=TRANSACTION_TYPES, verbose_name="操作類型")
    quantity_change = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="變動數量")
    new_confirmed_quantity = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="操作後總撥料數量")
    timestamp = models.DateTimeField(auto_now_add=True, verbose_name="操作時間")
    notes = models.CharField(max_length=255, blank=True, null=True, verbose_name="備註")

    class Meta:
        verbose_name = "訂單物料交易紀錄"
        verbose_name_plural = "訂單物料交易紀錄"
        ordering = ['-timestamp']

    def __str__(self):
        return f"{self.timestamp.strftime('%Y-%m-%d %H:%M')} - {self.work_order_material.material_number} - {self.get_transaction_type_display()}: {self.quantity_change}"

class WorkOrderMaterialImage(models.Model):
    requisition = models.ForeignKey(Requisition, on_delete=models.CASCADE, related_name='work_order_material_images', verbose_name="所屬撥料單", null=True, blank=True)
    process_type = models.ForeignKey(ProcessType, on_delete=models.SET_NULL, null=True, blank=True, verbose_name="投料點")
    image = models.ImageField(upload_to='work_order_material_images/', verbose_name="訂單物料圖片")
    uploaded_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, verbose_name="上傳人員")
    uploaded_at = models.DateTimeField(auto_now_add=True, verbose_name="上傳時間")

    class Meta:
        verbose_name = "訂單物料圖片"
        verbose_name_plural = "訂單物料圖片"
        ordering = ['-uploaded_at']

    def __str__(self):
        return f"圖片 for {self.requisition.order_number if self.requisition else 'N/A'} ({self.uploaded_at.strftime('%Y-%m-%d %H:%M')})"

class AutoUploadConfig(models.Model):
    UPLOAD_TYPES = [
        ('inventory', '庫存資料 (Inventory)'),
        ('order_model', '訂單機型 (Order Models)'),
        ('material_details', '物料明細 (Material Details)'),
        ('semi_finished', '半成品資料 (Semi-Finished)'),
        ('semi_finished_model_db', '半成品機型資料庫 (Semi-Finished Model DB)'),
    ]
    upload_type = models.CharField(max_length=50, choices=UPLOAD_TYPES, unique=True, verbose_name="上傳類型")
    file_path = models.CharField(max_length=255, verbose_name="檔案路徑", help_text="請輸入完整檔案路徑，例如 C:\\SAP\\inventory.xlsx")
    is_active = models.BooleanField(default=True, verbose_name="是否啟用")
    last_run = models.DateTimeField(null=True, blank=True, verbose_name="最後執行時間")
    last_status = models.CharField(max_length=255, blank=True, verbose_name="最後執行狀態")
    last_processed_mtime = models.FloatField(null=True, blank=True, verbose_name="最後處理檔案修改時間")
    priority = models.IntegerField(default=0, verbose_name="執行順序", help_text="數字越小越先執行")

    class Meta:
        verbose_name = "自動上傳設定"
        verbose_name_plural = "自動上傳設定"
        ordering = ['priority']

    def __str__(self):
        return f"{self.get_upload_type_display()} - {self.file_path}"

class MaterialProcessTypeRule(models.Model):
    material_prefix = models.CharField(max_length=100, verbose_name="物料前綴/號碼", db_index=True)
    machine_model_name = models.CharField(max_length=100, verbose_name="機型名稱", db_index=True)
    process_type_name = models.CharField(max_length=100, verbose_name="投料點名稱")
    parent_material_desc_keyword = models.CharField(max_length=100, blank=True, null=True, verbose_name="上層說明關鍵字", help_text="若填寫，則只有當上層物料說明包含此關鍵字時才套用")
    updated_at = models.DateTimeField(auto_now=True, verbose_name="最後更新時間")
    updated_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, verbose_name="更新人員")

    class Meta:
        verbose_name = "物料投料點規則 (學習紀錄)"
        verbose_name_plural = "物料投料點規則 (學習紀錄)"
        unique_together = ('material_prefix', 'machine_model_name', 'parent_material_desc_keyword')

    def __str__(self):
        return f"{self.material_prefix} + {self.machine_model_name} -> {self.process_type_name}"


class OperationProcessRule(models.Model):
    """作業說明對應投料點規則 - 學習型分類"""
    operation_description = models.CharField(max_length=255, unique=True, verbose_name="作業說明", db_index=True)
    process_type = models.CharField(max_length=50, verbose_name="投料點")
    created_at = models.DateTimeField(auto_now_add=True, verbose_name="建立時間")
    updated_at = models.DateTimeField(auto_now=True, verbose_name="最後更新時間")
    updated_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, verbose_name="設定人員")

    class Meta:
        verbose_name = "作業說明投料點規則"
        verbose_name_plural = "作業說明投料點規則"
        ordering = ['operation_description']

    def __str__(self):
        return f"{self.operation_description} -> {self.process_type}"


class SemiFinishedProcessType(models.Model):
    """半成品投料點 - 由主管動態管理"""
    name = models.CharField(max_length=100, unique=True, verbose_name="投料點名稱")
    color = models.CharField(max_length=20, default='#6366F1', verbose_name="顯示顏色")
    order = models.IntegerField(default=0, verbose_name="排序順序")
    is_active = models.BooleanField(default=True, verbose_name="是否啟用")
    created_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='created_semi_process_types', verbose_name="建立人員")
    created_at = models.DateTimeField(auto_now_add=True, verbose_name="建立時間")
    updated_at = models.DateTimeField(auto_now=True, verbose_name="最後更新時間")

    class Meta:
        verbose_name = "半成品投料點"
        verbose_name_plural = "半成品投料點"
        ordering = ['order', 'name']

    def __str__(self):
        return self.name


class UserProfile(models.Model):
    """使用者額外資訊，用於儲存申請角色等狀態"""
    user = models.OneToOneField(User, on_delete=models.CASCADE, related_name='profile', verbose_name="使用者")
    requested_role = models.CharField(max_length=50, blank=True, null=True, verbose_name="申請角色")
    
    class Meta:
        verbose_name = "使用者設定檔"
        verbose_name_plural = "使用者設定檔"

    def __str__(self):
        return f"{self.user.username} 的設定檔"


class RequisitionViewPermission(models.Model):
    """申請單查看授權 - 允許某使用者查看並操作另一使用者的申請單"""
    owner = models.ForeignKey(User, on_delete=models.CASCADE,
        related_name='view_permissions_granted', verbose_name="申請單擁有者")
    viewer = models.ForeignKey(User, on_delete=models.CASCADE,
        related_name='view_permissions_received', verbose_name="被授權查看者")

    class Meta:
        verbose_name = "申請單查看授權"
        verbose_name_plural = "申請單查看授權"
        unique_together = ('owner', 'viewer')

    def __str__(self):
        return f"{self.viewer.username} 可查看 {self.owner.username} 的申請單"

