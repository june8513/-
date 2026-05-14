"""
共用工具函數
跨 App 使用的通用功能
"""
from django.contrib.auth.models import User


def get_sap_user():
    """
    取得或建立名為 SAP已撥料 的虛擬使用者
    用於標記 SAP 系統自動撥料的操作人員
    """
    sap_user, created = User.objects.get_or_create(
        username='sap_system',
        defaults={
            'first_name': 'SAP',
            'last_name': '已撥料',
            'is_active': False  # 系統帳號不允許登入
        }
    )
    return sap_user
