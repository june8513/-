"""
開發環境設定
"""
from .base import *

DEBUG = True

ALLOWED_HOSTS = ['*']

CSRF_TRUSTED_ORIGINS = [
    'https://973a749b995d.ngrok-free.app',
    'https://d6992b01684b.ngrok-free.app',
    'http://192.168.6.137',
    'http://192.168.6.137:8000',
    'http://192.168.6.137:8001',
    'http://192.168.6.137:8002',
]
