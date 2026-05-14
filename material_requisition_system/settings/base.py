"""
Django base settings - 所有環境共用的設定
"""
import os
from pathlib import Path

# Build paths inside the project like this: BASE_DIR / 'subdir'.
# 注意：settings 已移至子目錄，需要多往上一層
BASE_DIR = Path(__file__).resolve().parent.parent.parent


# ─── 安全性 ─────────────────────────────────────────

SECRET_KEY = os.environ.get(
    'DJANGO_SECRET_KEY',
    'django-insecure-n=3!_=t)j8m-c@kx_i!8z7^)&e0bmv-+o5=u+3+mgt_v425+%n'
)

# ─── 表單 ───────────────────────────────────────────

DATA_UPLOAD_MAX_NUMBER_FIELDS = 10000

# ─── 外部 API Key ──────────────────────────────────

API_KEY = os.environ.get('YOUR_API_KEY_NAME', 'default_api_key_if_not_set')

# ─── 應用程式 ───────────────────────────────────────

INSTALLED_APPS = [
    'django.contrib.admin',
    'django.contrib.auth',
    'django.contrib.contenttypes',
    'django.contrib.sessions',
    'django.contrib.messages',
    'django.contrib.staticfiles',
    'django_extensions',
    'requisitions',
    'peer_requests',
    'widget_tweaks',
    'sslserver',
    'inventory',
    'specifications',
    'core',
]

MIDDLEWARE = [
    'django.middleware.security.SecurityMiddleware',
    'whitenoise.middleware.WhiteNoiseMiddleware',
    'django.contrib.sessions.middleware.SessionMiddleware',
    'django.middleware.common.CommonMiddleware',
    'django.middleware.csrf.CsrfViewMiddleware',
    'django.contrib.auth.middleware.AuthenticationMiddleware',
    'django.contrib.messages.middleware.MessageMiddleware',
    'django.middleware.clickjacking.XFrameOptionsMiddleware',
]

ROOT_URLCONF = 'material_requisition_system.urls'

TEMPLATES = [
    {
        'BACKEND': 'django.template.backends.django.DjangoTemplates',
        'DIRS': [BASE_DIR / 'templates'],
        'APP_DIRS': True,
        'OPTIONS': {
            'context_processors': [
                'django.template.context_processors.request',
                'django.contrib.auth.context_processors.auth',
                'django.contrib.messages.context_processors.messages',
                'core.context_processors.role_context',
            ],
        },
    },
]

WSGI_APPLICATION = 'material_requisition_system.wsgi.application'

# ─── 資料庫 ─────────────────────────────────────────

DATABASES = {
    'default': {
        'ENGINE': 'django.db.backends.sqlite3',
        'NAME': BASE_DIR / 'db.sqlite3',
        'OPTIONS': {
            'timeout': 60,
            'transaction_mode': 'IMMEDIATE',
        }
    }
}

# ─── 密碼驗證 ───────────────────────────────────────

AUTH_PASSWORD_VALIDATORS = []

# ─── 國際化 ─────────────────────────────────────────

LANGUAGE_CODE = 'en-us'
TIME_ZONE = 'Asia/Taipei'
USE_I18N = True
USE_TZ = True

# ─── Session ────────────────────────────────────────

SESSION_EXPIRE_AT_BROWSER_CLOSE = True

# ─── 靜態檔案 ───────────────────────────────────────

STATIC_URL = 'static/'
STATIC_ROOT = os.path.join(BASE_DIR, 'staticfiles')
STATICFILES_DIRS = [BASE_DIR / 'static']
STATICFILES_STORAGE = 'whitenoise.storage.CompressedManifestStaticFilesStorage'

# ─── Media 檔案 ─────────────────────────────────────

MEDIA_URL = '/media/'
MEDIA_ROOT = os.path.join(BASE_DIR, 'media')

# ─── 預設主鍵 ───────────────────────────────────────

DEFAULT_AUTO_FIELD = 'django.db.models.BigAutoField'

# ─── 認證 ───────────────────────────────────────────

LOGIN_URL = '/requisitions/login/'
LOGIN_REDIRECT_URL = '/'

# ─── 外部服務 ───────────────────────────────────────

SHORTAGE_NOTIFICATION_URL = os.environ.get(
    'SHORTAGE_NOTIFICATION_URL',
    'https://httpbin.org/post'
)
