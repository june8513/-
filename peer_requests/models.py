from django.db import models
from django.contrib.auth.models import User

class PeerRequest(models.Model):
    STATUS_CHOICES = [
        ('pending', '待處理'),
        ('processing', '處理中'),
        ('shipped', '已撥料'),
        ('closed', '已結案'),
    ]

    applicant = models.ForeignKey(User, on_delete=models.CASCADE, related_name='peer_requests_made', verbose_name="申請人")
    recipient = models.ForeignKey(User, on_delete=models.CASCADE, related_name='peer_requests_received', verbose_name="收件人")
    cc_users = models.ManyToManyField(User, related_name='peer_requests_cced', blank=True, verbose_name="副本通知(CC)")

    
    # 需求部分 (第2、3欄)
    description = models.TextField(verbose_name="需求內容")
    request_photo = models.ImageField(upload_to='peer_requests/requests/', null=True, blank=True, verbose_name="需求照片")
    request_date = models.DateField(verbose_name="需求日期")
    
    # 回覆部分 (第4欄)
    expected_delivery_date = models.DateField(null=True, blank=True, verbose_name="預計完成日期")
    delivery_date = models.DateField(null=True, blank=True, verbose_name="實際撥料日期")
    delivery_photo = models.ImageField(upload_to='peer_requests/deliveries/', null=True, blank=True, verbose_name="撥料照片")
    delivery_reply = models.TextField(null=True, blank=True, verbose_name="收件人回覆文字")
    
    # 狀態與結案 (第5欄)
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='pending', verbose_name="狀態")
    closed_at = models.DateTimeField(null=True, blank=True, verbose_name="結案時間")
    
    created_at = models.DateTimeField(auto_now_add=True, verbose_name="建立時間")
    updated_at = models.DateTimeField(auto_now=True, verbose_name="更新時間")

    class Meta:
        verbose_name = "物料申請"
        verbose_name_plural = "物料申請"
        ordering = ['-created_at']

    def __str__(self):
        return f"Request from {self.applicant} to {self.recipient} - {self.status}"

class PeerRequestTemplate(models.Model):
    user = models.ForeignKey(User, on_delete=models.CASCADE, related_name='peer_request_templates', verbose_name="擁有者")
    name = models.CharField(max_length=50, verbose_name="模板名稱")
    recipient = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='+', verbose_name="預設收件人")
    cc_users = models.ManyToManyField(User, blank=True, related_name='+', verbose_name="預設副本(CC)")
    description = models.TextField(verbose_name="預設需求內容")
    created_at = models.DateTimeField(auto_now_add=True, verbose_name="建立時間")

    class Meta:
        verbose_name = "物料申請模板"
        verbose_name_plural = "物料申請模板"
        ordering = ['-created_at']

    def __str__(self):
        return f"Template '{self.name}' for {self.user.username}"
