"""
Django settings for maximo_project project.

Generated by 'django-admin startproject' using Django 5.1.

For more information on this file, see
https://docs.djangoproject.com/en/5.1/topics/settings/

For the full list of settings and their values, see
https://docs.djangoproject.com/en/5.1/ref/settings/
"""

from pathlib import Path
import os
import sys
from logging.handlers import RotatingFileHandler
from concurrent_log_handler import ConcurrentRotatingFileHandler

# ==============================================================================
# PATHS
# ==============================================================================
# Build paths inside the project like this: BASE_DIR / 'subdir'.
BASE_DIR = Path(__file__).resolve().parent.parent


# Quick-start development settings - unsuitable for production
# See https://docs.djangoproject.com/en/5.1/howto/deployment/checklist/

# ==============================================================================
# SECURITY SETTINGS
# ==============================================================================
# SECURITY WARNING: keep the secret key used in production secret!
SECRET_KEY = 'django-insecure-8qr^6l&nb!g6g22xu(^6h@wb#hc$54e@qr76(@x*npdvhh#&!u'

# SECURITY WARNING: don't run with debug turned on in production!
DEBUG = False    # ควรตั้งค่าเป็น False ใน production
ALLOWED_HOSTS = ['localhost', '127.0.0.1', '8dbb-161-246-199-200.ngrok-free.app']
CSRF_TRUSTED_ORIGINS = ['https://8dbb-161-246-199-200.ngrok-free.app']
# CSRF_COOKIE_SECURE = True   # ส่ง CSRF Cookie เฉพาะผ่าน HTTPS
# SESSION_COOKIE_SECURE = True  # ใช้ Secure Cookie (เฉพาะ HTTPS)
PORT = int(os.environ.get("PORT", 8000))

# ==============================================================================
# APPLICATION DEFINITION
# ==============================================================================
INSTALLED_APPS = [
    'django.contrib.admin',
    'django.contrib.auth',
    'django.contrib.contenttypes',
    'django.contrib.sessions',
    'django.contrib.messages',
    'django.contrib.staticfiles',
    'maximo_app',
    'background_task',
    'compressor',
    'widget_tweaks',
]

MIDDLEWARE = [
    'django.middleware.security.SecurityMiddleware',
    'django.contrib.sessions.middleware.SessionMiddleware',
    'django.middleware.common.CommonMiddleware',
    'django.middleware.csrf.CsrfViewMiddleware',
    'django.contrib.auth.middleware.AuthenticationMiddleware',
    'django.contrib.messages.middleware.MessageMiddleware',
    'django.middleware.clickjacking.XFrameOptionsMiddleware',
    'whitenoise.middleware.WhiteNoiseMiddleware',
]

ROOT_URLCONF = 'maximo_project.urls'

TEMPLATES = [
    {
        'BACKEND': 'django.template.backends.django.DjangoTemplates',
        'DIRS': [os.path.join(BASE_DIR, 'maximo_app', 'templates')],
        'APP_DIRS': True,
        'OPTIONS': {
            'context_processors': [
                'django.template.context_processors.debug',
                'django.template.context_processors.request',
                'django.contrib.auth.context_processors.auth',
                'django.contrib.messages.context_processors.messages',
            ],
        },
    },
]

WSGI_APPLICATION = 'maximo_project.wsgi.application'

# ==============================================================================
# DATABASE CONFIGURATION
# ==============================================================================
# https://docs.djangoproject.com/en/5.1/ref/settings/#databases

DATABASES = {
    'default': {
        'ENGINE': 'django.db.backends.sqlite3',
        'NAME': BASE_DIR / 'db.sqlite3',
    }
}

# ==============================================================================
# PASSWORD VALIDATION
# ==============================================================================
# https://docs.djangoproject.com/en/5.1/ref/settings/#auth-password-validators

AUTH_PASSWORD_VALIDATORS = [
    {
        'NAME': 'django.contrib.auth.password_validation.UserAttributeSimilarityValidator',
    },
    {
        'NAME': 'django.contrib.auth.password_validation.MinimumLengthValidator',
    },
    {
        'NAME': 'django.contrib.auth.password_validation.CommonPasswordValidator',
    },
    {
        'NAME': 'django.contrib.auth.password_validation.NumericPasswordValidator',
    },
]

# ==============================================================================
# INTERNATIONALIZATION
# ==============================================================================
# https://docs.djangoproject.com/en/5.1/topics/i18n/

LANGUAGE_CODE = 'en-us'
TIME_ZONE = 'UTC'
USE_I18N = True
USE_TZ = True

# ==============================================================================
# STATIC FILES (CSS, JavaScript, Images)
# ==============================================================================
# https://docs.djangoproject.com/en/5.1/howto/static-files/

STATIC_URL = '/static/'

# STATIC_ROOT ใช้สำหรับ production environment หลังจากรันคำสั่ง collectstatic
STATIC_ROOT = os.path.join(BASE_DIR, 'staticfiles')

# STATICFILES_DIRS ใช้ใน development environment เพื่อบอกว่า static files อยู่ที่ไหน
STATICFILES_DIRS = [
    BASE_DIR / "static",
]

# Default primary key field type
# https://docs.djangoproject.com/en/5.1/ref/settings/#default-auto-field

DEFAULT_AUTO_FIELD = 'django.db.models.BigAutoField'

# ==============================================================================
# MEDIA FILES
# ==============================================================================
MEDIA_URL = '/media/'
MEDIA_ROOT = BASE_DIR / 'media'

# ==============================================================================
# LOGGING CONFIGURATION
# ==============================================================================
LOGS_DIR = os.path.join(BASE_DIR, 'logs')
if not os.path.exists(LOGS_DIR):
    os.makedirs(LOGS_DIR)

LOGGING = {
    'version': 1,
    'disable_existing_loggers': False,
    'formatters': {
        'verbose_default': {
            'format': '[%(asctime)s] %(levelname)s [%(name)s:%(lineno)s] %(message)s',
            'datefmt': '%d/%b/%Y %H:%M:%S',
        },
        'simple': {
            'format': '{levelname} {message}',
            'style': '{',
        },
    },
    'handlers': {
        'file_info': {
            'level': 'INFO',
            'class': 'logging.handlers.RotatingFileHandler',
            'filename': os.path.join(LOGS_DIR, 'info.log'),
            'formatter': 'verbose_default',
            'maxBytes': 10 * 1024 * 1024,
            'backupCount': 10,
            'encoding': 'utf-8',
        },
        'file_warning': {
            'level': 'WARNING',
            'class': 'logging.handlers.RotatingFileHandler',
            'filename': os.path.join(LOGS_DIR, 'warning.log'),
            'formatter': 'verbose_default',
            'maxBytes': 10 * 1024 * 1024,
            'backupCount': 10,
            'encoding': 'utf-8',
        },
        'file_error': {
            'level': 'ERROR',
            'class': 'logging.handlers.RotatingFileHandler',
            'filename': os.path.join(LOGS_DIR, 'error.log'),
            'formatter': 'verbose_default',
            'maxBytes': 10 * 1024 * 1024,
            'backupCount': 10,
            'encoding': 'utf-8',
            
        },
        'file_critical': {
            'level': 'CRITICAL',
            'class': 'logging.handlers.RotatingFileHandler',
            'filename': os.path.join(LOGS_DIR, 'critical.log'),
            'formatter': 'verbose_default',
            'maxBytes': 10 * 1024 * 1024,
            'backupCount': 10,
            'encoding': 'utf-8',
        },
        'file_debug': {
            'level': 'DEBUG',
            'class': 'concurrent_log_handler.ConcurrentRotatingFileHandler',
            'filename': os.path.join(LOGS_DIR, 'debug.log'),
            'formatter': 'verbose_default',
            'maxBytes': 10 * 1024 * 1024,
            'backupCount': 10,
            'encoding': 'utf-8',
        },
        'console': {
            'class': 'logging.StreamHandler',
            'formatter': 'simple',
            'level': 'DEBUG',
            'stream': sys.stdout,
        },
    },
    'loggers': {
        'django': {
            'handlers': ['file_info', 'file_error', 'file_warning', 'file_critical', 'file_debug', 'console'],  # กำหนดว่าข้อความ log จะถูกจัดการอย่างไรหรือจะถูกส่งไปที่ไหน Console: แสดง log ใน terminal หรือ console ,File: บันทึก log ลงไฟล์
            'level': 'DEBUG',   # ระดับความสำคัญของข้อความที่จะถูกบันทึก
            'propagate': True,
        },
        'maximo_app': {
            'handlers': ['file_info', 'file_error', 'file_warning', 'file_critical', 'file_debug', 'console'],
            'level': 'DEBUG',
            'propagate': True,
        },
    },
}

# ==============================================================================
# SESSION CONFIGURATION
# ==============================================================================
SESSION_EXPIRE_AT_BROWSER_CLOSE = True  # เซสชันจะหมดอายุเมื่อปิดเบราว์เซอร์
SESSION_COOKIE_AGE = 3600  # เซสชันหมดอายุหลังจาก 1 ชม.
SESSION_ENGINE = 'django.contrib.sessions.backends.db'

# ==============================================================================
# COMPRESSOR CONFIGURATION
# ==============================================================================
COMPRESS_ROOT = BASE_DIR / 'static'
COMPRESS_ENABLED = True
STATICFILES_FINDERS = (
    'django.contrib.staticfiles.finders.FileSystemFinder',
    'django.contrib.staticfiles.finders.AppDirectoriesFinder',
    'compressor.finders.CompressorFinder',
)

# ==============================================================================
# CUSTOM SETTINGS
# ==============================================================================
TEMP_DIR = BASE_DIR / 'temp'
FILE_AGE_LIMIT = 3600
STATICFILES_STORAGE = 'whitenoise.storage.CompressedManifestStaticFilesStorage'