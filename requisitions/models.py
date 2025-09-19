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

    class Meta:
        verbose_name = "投料點"
        verbose_name_plural = "投料點"
        unique_together = ('name', 'machine_model')
        ordering = ['machine_model', 'name']

    def __str__(self):
        return self.name

class Requisition(models.Model):
    STATUS_CHOICES = [
        ('pending', '待處理'),
        ('materials_confirmed', '物料已確認'),
        ('completed', '已處理'),
    ]

    order_number = models.CharField(max_length=100, verbose_name="訂單單號")
    applicant = models.ForeignKey(User, on_delete=models.CASCADE, related_name='requisitions_applied', verbose_name="申請人")
    request_date = models.DateField(verbose_name="需求日期", db_index=True)
    process_type = models.CharField(max_length=100, verbose_name="需求流程", db_index=True, null=True, blank=True)
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='pending', verbose_name="狀態", db_index=True)
    dispatch_performed = models.BooleanField(default=False, verbose_name="已執行撥料")
    
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
            UniqueConstraint(fields=['order_number', 'process_type'], name='unique_order_per_process')
        ]

    def __str__(self):
        return f"撥料申請單: {self.order_number} ({self.process_type}) - {self.applicant.username}"

class MaterialListVersion(models.Model):
    requisition = models.ForeignKey(Requisition, on_delete=models.CASCADE, related_name='material_versions', verbose_name="所屬撥料單")
    uploaded_at = models.DateTimeField(auto_now_add=True, verbose_name="上傳時間")
    uploaded_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, verbose_name="上傳人員")

    class Meta:
        verbose_name = "物料清單版本"
        verbose_name_plural = "物料清單版本"
        ordering = ['-uploaded_at']

    def __str__(self):
        return f"{self.requisition.order_number} - 物料版本 ({self.uploaded_at.strftime('%Y-%m-%d %H:%M')})"

class WorkOrderMaterial(models.Model):
    machine_model = models.ForeignKey(MachineModel, on_delete=models.CASCADE, verbose_name="機型", related_name="work_order_materials", null=True, blank=True)
    order_number = models.CharField(max_length=100, db_index=True, verbose_name="訂單單號")
    material_number = models.CharField(max_length=100, db_index=True, verbose_name="物料", null=True, blank=True)
    item_name = models.CharField(max_length=255, verbose_name="品名", null=True, blank=True)
    required_quantity = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="需求數量")
    process_type = models.ForeignKey(ProcessType, on_delete=models.SET_NULL, null=True, blank=True, verbose_name="投料點")
    
    confirmed_quantity = models.DecimalField(max_digits=10, decimal_places=2, null=True, blank=True, verbose_name="已撥料數量")
    is_signed_off = models.BooleanField(default=False, verbose_name="已簽收")
    is_active = models.BooleanField(default=True, verbose_name="是否啟用") # New field

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "訂單主物料清單"
        verbose_name_plural = "訂單主物料清單"
        ordering = ['order_number', 'process_type', 'material_number']
        

    def __str__(self):
        return f"{self.order_number} - {self.machine_model} - {self.material_number} ({self.item_name})"

class RequisitionItem(models.Model):
    material_list_version = models.ForeignKey(MaterialListVersion, on_delete=models.CASCADE, null=True, blank=True, related_name='items', verbose_name="物料清單版本")
    source_material = models.ForeignKey(WorkOrderMaterial, on_delete=models.SET_NULL, null=True, blank=True, related_name='requisition_items', verbose_name="來源主物料")
    order_number = models.CharField(max_length=100, verbose_name="訂單單號")
    material_number = models.CharField(max_length=100, verbose_name="物料", db_index=True)
    item_name = models.CharField(max_length=255, verbose_name="品名")
    required_quantity = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="需求數量")
    stock_quantity = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="庫存數量")
    storage_bin = models.CharField(max_length=100, blank=True, null=True, verbose_name="儲格")
    
    confirmed_quantity = models.DecimalField(max_digits=10, decimal_places=2, null=True, blank=True, verbose_name="確認撥料數量")
    is_signed_off = models.BooleanField(default=False, verbose_name="最終簽收已確認")

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
