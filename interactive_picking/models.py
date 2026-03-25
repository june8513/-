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
    floor = models.IntegerField('樓層', default=1, help_text='預設為 1 樓')
    description = models.CharField('位置描述', max_length=200, blank=True)

    class Meta:
        verbose_name = '倉儲節點'
        verbose_name_plural = '倉儲節點'

    def __str__(self):
        return f"{self.name} ({self.floor}F - {self.x_coordinate}, {self.y_coordinate})"

    def distance_to(self, other_location):
        """計算與另一個節點的直線距離"""
        if not other_location:
            return float('inf')
        
        # 同樓層正常計算，若跨樓層則給予一個高度懲罰值 (例如每層樓 1000 距離)
        floor_penalty = abs(getattr(self, 'floor', 1) - getattr(other_location, 'floor', 1)) * 1000
        
        return math.sqrt(
            (self.x_coordinate - other_location.x_coordinate)**2 +
            (self.y_coordinate - other_location.y_coordinate)**2
        ) + floor_penalty



