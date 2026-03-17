import math
from django.db import models
from django.utils import timezone

class WarehouseLocation(models.fields.CharField):
    # This is not a real django model, but a placeholder for now since we just created the file
    pass

class WarehouseLocation(models.Model):
    """
    倉儲節點：定義每個儲位的實體座標，用於路徑最佳化計算
    """
    name = models.CharField('儲位代號', max_length=50, unique=True, help_text='例如：A-01, B-02-03')
    x_coordinate = models.FloatField('X 座標', default=0.0, help_text='空間相對 X 座標')
    y_coordinate = models.FloatField('Y 座標', default=0.0, help_text='空間相對 Y 座標')
    description = models.CharField('位置描述', max_length=200, blank=True)

    class Meta:
        verbose_name = '倉儲節點'
        verbose_name_plural = '倉儲節點'

    def __str__(self):
        return f"{self.name} ({self.x_coordinate}, {self.y_coordinate})"

    def distance_to(self, other_location):
        """計算與另一個節點的直線距離"""
        if not other_location:
            return float('inf')
        return math.sqrt(
            (self.x_coordinate - other_location.x_coordinate)**2 +
            (self.y_coordinate - other_location.y_coordinate)**2
        )


class MockPickingTask(models.Model):
    """
    模擬揀料單：一次出車/揀料作業的批次單
    """
    task_number = models.CharField('任務單號', max_length=50, unique=True)
    created_at = models.DateTimeField('建立時間', default=timezone.now)
    is_completed = models.BooleanField('是否完成', default=False)

    class Meta:
        verbose_name = '模擬揀料單'
        verbose_name_plural = '模擬揀料單'

    def __str__(self):
        return f"任務: {self.task_number}"


class MockPickingItem(models.Model):
    """
    模擬揀料項目：對應單一欲揀取的物料與其儲位
    """
    STATUS_CHOICES = [
        ('pending', '待拿取'),
        ('picked', '已撥料'),
        ('shortage', '缺料/跳過'),
    ]

    task = models.ForeignKey(MockPickingTask, on_delete=models.CASCADE, related_name='items')
    material_name = models.CharField('物料名稱', max_length=100)
    quantity_required = models.PositiveIntegerField('應拿數量')
    location = models.ForeignKey(WarehouseLocation, on_delete=models.SET_NULL, null=True, related_name='picking_items')
    status = models.CharField('當前狀態', max_length=20, choices=STATUS_CHOICES, default='pending')
    picked_at = models.DateTimeField('處理時間', null=True, blank=True)

    class Meta:
        verbose_name = '模擬揀料項目'
        verbose_name_plural = '模擬揀料項目'

    def __str__(self):
        return f"{self.material_name} ({self.quantity_required}件) [{self.get_status_display()}]"
