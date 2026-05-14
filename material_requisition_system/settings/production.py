"""
正式環境設定
"""
import os
from .base import *

DEBUG = False

# 從環境變數讀取 ALLOWED_HOSTS，如果沒有則預設為 *（建議在正式環境中明確指定）
allowed_hosts_env = os.environ.get('DJANGO_ALLOWED_HOSTS', '*')
ALLOWED_HOSTS = [host.strip() for host in allowed_hosts_env.split(',') if host.strip()]

# 從環境變數讀取 CSRF_TRUSTED_ORIGINS
csrf_origins_env = os.environ.get('DJANGO_CSRF_TRUSTED_ORIGINS', '')
if csrf_origins_env:
    CSRF_TRUSTED_ORIGINS = [origin.strip() for origin in csrf_origins_env.split(',') if origin.strip()]

# 安全性設定 (可選，視反向代理設定而定)
# SECURE_PROXY_SSL_HEADER = ('HTTP_X_FORWARDED_PROTO', 'https')
# SESSION_COOKIE_SECURE = True
# CSRF_COOKIE_SECURE = True
