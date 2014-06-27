# -*- coding: utf-8 -*-
"""
Django settings for daiweb project.

For more information on this file, see
https://docs.djangoproject.com/en/1.6/topics/settings/

For the full list of settings and their values, see
https://docs.djangoproject.com/en/1.6/ref/settings/
"""
# Importacion de librerias
import os
import sys
from django.conf.global_settings import TEMPLATE_CONTEXT_PROCESSORS as TCP
from django.core.urlresolvers import reverse_lazy

# configuracion para django-suit
TEMPLATE_CONTEXT_PROCESSORS = TCP + (
    'django.contrib.auth.context_processors.auth',
    'django.core.context_processors.i18n',
    'django.core.context_processors.request',
    'django.core.context_processors.media',
    'django.core.context_processors.static',
    #'cms.context_processors.media',
    'sekizai.context_processors.sekizai',
)
# estableciendo la ruta del directorio

gettext = lambda s: s
PROJECT_PATH = os.path.split(os.path.abspath(os.path.dirname(__file__)))[0]

# Build paths inside the project like this: os.path.join(BASE_DIR, ...)

BASE_DIR = os.path.dirname(os.path.dirname(__file__))


# Quick-start development settings - unsuitable for production
# See https://docs.djangoproject.com/en/1.6/howto/deployment/checklist/

# SECURITY WARNING: keep the secret key used in production secret!
SECRET_KEY = '*cp+(b^9^9ukx6z#4c(ky1mv-y)heb+pfcaabnj)55!aaf28*@'

# SECURITY WARNING: don't run with debug turned on in production!
DEBUG = True

TEMPLATE_DEBUG = True

ALLOWED_HOSTS = []


# Application definition

INSTALLED_APPS = (
    'suit',
    #'django_admin_bootstrapped',
    'django.contrib.admin',
    'django.contrib.auth',
    'django.contrib.contenttypes',
    'django.contrib.sessions',
    'django.contrib.messages',
    'django.contrib.staticfiles',
    'django.contrib.sites',
    'webportal',
    'menus',
    'cms',
    'mptt',
    'south',
    'sekizai',
    'tracking',
    'tickets',
    'sitetree',
    #'control',
    'cfdi',
    #'cms.plugins.flash',
    #'cms.plugins.googlemap',
    #'cms.plugins.link',
    #'cms.plugins.text-ckeditor',
    'filer',
    'easy_thumbnails',
    'previos',
    #'cmsplugin_filer_file',
    #'cmsplugin_filer_folder',
    #'cmsplugin_filer_image',
    #'cmsplugin_filer_teaser',
    #'cmsplugin_filer_video',
)

MIDDLEWARE_CLASSES = (
    'django.contrib.sessions.middleware.SessionMiddleware',
    'django.middleware.csrf.CsrfViewMiddleware',
    'django.contrib.auth.middleware.AuthenticationMiddleware',
    'django.contrib.messages.middleware.MessageMiddleware',
    'django.middleware.locale.LocaleMiddleware',
    'django.middleware.doc.XViewMiddleware',
    'django.middleware.common.CommonMiddleware',
    #'cms.middleware.page.CurrentPageMiddleware',
    #'cms.middleware.user.CurrentUserMiddleware',
    #'cms.middleware.toolbar.ToolbarMiddleware',
    #'cms.middleware.language.LanguageCookieMiddleware',
)

ROOT_URLCONF = 'daiweb.urls'

WSGI_APPLICATION = 'daiweb.wsgi.application'
CSRF_COOKIE_SECURE = False

# Database
# https://docs.djangoproject.com/en/1.6/ref/settings/#databases

DATABASES = {
    'default': {
        'ENGINE': 'django.db.backends.sqlite3',
        'NAME': os.path.join(BASE_DIR, 'daiweb.sqlite3'),
    }
    #'default': {
    #    'ENGINE':'django.db.backends.postgresql_psycopg2',
    #    'NAME': 'dai',
    #    'USER': 'admindai',
    #    'PASSWORD': 'dai123',
    #    'HOST': '10.66.10.130',
    #    'PORT': '5432',
    #} 
}

# Internationalization
# https://docs.djangoproject.com/en/1.6/topics/i18n/

LANGUAGE_CODE = 'es-MX'

TIME_ZONE = 'America/Monterrey'
    
USE_I18N = True

USE_L10N = True

USE_TZ = True


# Static files (CSS, JavaScript, Images)
# https://docs.djangoproject.com/en/1.6/howto/static-files/

# Variables para la redireccion de login y logout
# Variables para la autenticacion de usuario
LOGIN_URL = reverse_lazy('entrar')
LOGOUT_URL = reverse_lazy('salir')
#Redireccion principal
LOGIN_REDIRECT_URL = reverse_lazy('entrar')

STATIC_ROOT = os.path.join(PROJECT_PATH, "static")
STATIC_URL = "/static/"
ADMIN_MEDIA_PREFIX = '/static/admin/'
 # Para multimedia
MEDIA_ROOT = os.path.join(PROJECT_PATH, "media")
MEDIA_URL = "/media/"

# seccion de las plantillas
TEMPLATE_DIRS = (
    # The docs say it should be absolute path: PROJECT_PATH is precisely one.
    # Life is wonderful!
    os.path.join(PROJECT_PATH, "plantillas"),
    
)

# Idiomas
LANGUAGES = [
    ('en', 'English'),
    ('es', 'Spanish'),
    ('es-MX','Spanish (Mexico)'),
]

SITE_ID = 1