from django.db import models
from inventory.models import Material

class MaterialSpecification(models.Model):
    material = models.OneToOneField(Material, on_delete=models.CASCADE, related_name='specification', verbose_name="物料")
    size = models.CharField(max_length=100, blank=True, null=True, verbose_name="大小")
    weight = models.DecimalField(max_digits=10, decimal_places=2, blank=True, null=True, verbose_name="重量")
    detailed_description = models.TextField(blank=True, null=True, verbose_name="詳細說明")
    image = models.ImageField(upload_to='material_spec_images/', blank=True, null=True, verbose_name="圖片")

    def __str__(self):
        return f"Specification for {self.material.material_code}"