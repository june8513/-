from django.db import models
from django.contrib.auth.models import User

class StorageLocation(models.Model):
    name = models.CharField(max_length=100, unique=True, verbose_name="儲存地點")

    def __str__(self):
        return self.name

# 1. Master Inventory List Model
class Material(models.Model):
    location = models.ForeignKey(StorageLocation, on_delete=models.PROTECT, verbose_name="庫位")
    bin = models.CharField(max_length=100, verbose_name="儲格")
    material_code = models.CharField(max_length=100, unique=True, verbose_name="物料")
    material_description = models.CharField(max_length=200, verbose_name="物料說明")
    system_quantity = models.IntegerField(verbose_name="系統庫存數量")
    last_counted_date = models.DateTimeField(null=True, blank=True, verbose_name="上次盤點日期")
    latest_counted_quantity = models.IntegerField(null=True, blank=True, verbose_name="最新盤點數量") # Keep this field
    last_counted_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, verbose_name="盤點人員")

    class Meta:
        verbose_name = "主物料"
        verbose_name_plural = "主庫存清單"

    @property
    def current_difference(self):
        if self.latest_counted_quantity is not None:
            return self.latest_counted_quantity - self.system_quantity
        return None

    def __str__(self):
        return self.material_code

# 2. Stocktake Header Model
class Stocktake(models.Model):
    STATUS_CHOICES = [
        ('進行中', '進行中'),
        ('已完成', '已完成'),
    ]
    stocktake_id = models.CharField(max_length=50, unique=True, verbose_name="盤點單號")
    name = models.CharField(max_length=100, blank=True, null=True, verbose_name="盤點單名稱")
    created_by = models.ForeignKey(User, on_delete=models.PROTECT, verbose_name="建立人員")
    created_at = models.DateTimeField(auto_now_add=True, verbose_name="建立日期")
    status = models.CharField(max_length=10, choices=STATUS_CHOICES, default='進行中', verbose_name="狀態")
    # notes = models.TextField(blank=True, null=True, verbose_name="筆記") # Removed field

    class Meta:
        verbose_name = "盤點單"
        verbose_name_plural = "盤點單列表"

    def __str__(self):
        return self.name if self.name else self.stocktake_id # Use name if available

# 3. Stocktake Item Model (linking the two above)
class StocktakeItem(models.Model):
    STATUS_CHOICES = [
        ('待盤點', '待盤點'),
        ('已盤點', '已盤點'),
    ]
    stocktake = models.ForeignKey(Stocktake, related_name='items', on_delete=models.CASCADE, verbose_name="盤點單")
    material = models.ForeignKey(Material, on_delete=models.PROTECT, verbose_name="物料")
    system_quantity_on_record = models.IntegerField(verbose_name="當時系統數量")
    counted_quantity = models.IntegerField(null=True, blank=True, verbose_name="實際盤點數量")
    status = models.CharField(max_length=10, choices=STATUS_CHOICES, default='待盤點', verbose_name="狀態")

    class Meta:
        verbose_name = "盤點品項"
        verbose_name_plural = "盤點品項列表"

    @property
    def difference(self):
        if self.counted_quantity is not None:
            return self.counted_quantity - self.system_quantity_on_record
        return None

    def __str__(self):
        return f"{self.stocktake.stocktake_id} - {self.material.material_code}"

# 4. Material Transaction History Model
class MaterialTransaction(models.Model):
    TRANSACTION_TYPES = [
        ('ALLOCATION', '撥料'),
        ('RETURN', '退料'),
        ('MANUAL_UPDATE', '手動修改'),
        ('INITIAL_IMPORT', '初始匯入'),
    ]

    material = models.ForeignKey(Material, on_delete=models.CASCADE, related_name='transactions', verbose_name="物料")
    user = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, verbose_name="操作人員")
    transaction_type = models.CharField(max_length=20, choices=TRANSACTION_TYPES, verbose_name="操作類型")
    quantity_change = models.IntegerField(verbose_name="變動數量")
    new_system_quantity = models.IntegerField(verbose_name="操作後總量")
    timestamp = models.DateTimeField(auto_now_add=True, verbose_name="操作時間")
    notes = models.CharField(max_length=255, blank=True, null=True, verbose_name="備註")

    class Meta:
        verbose_name = "物料交易紀錄"
        verbose_name_plural = "物料交易紀錄"
        ordering = ['-timestamp']

    def __str__(self):
        return f"{self.timestamp.strftime('%Y-%m-%d %H:%M')} - {self.material.material_code} - {self.get_transaction_type_display()}: {self.quantity_change}"


class MaterialImage(models.Model):
    material = models.ForeignKey(Material, on_delete=models.CASCADE, related_name='images', verbose_name="所屬物料")
    image = models.ImageField(upload_to='material_images/', verbose_name="物料圖片")
    uploaded_at = models.DateTimeField(auto_now_add=True, verbose_name="上傳時間")
    uploaded_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, verbose_name="上傳人員")

    class Meta:
        verbose_name = "物料圖片"
        verbose_name_plural = "物料圖片"
        ordering = ['-uploaded_at']

    def __str__(self):
        return f"圖片 for {self.material.material_code} ({self.uploaded_at.strftime('%Y-%m-%d %H:%M')})"