"""
URL configuration for webger23 project.

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/4.2/topics/http/urls/
Examples:
Function views
    1. Add an import:  from my_app import views
    2. Add a URL to urlpatterns:  path('', views.home, name='home')
Class-based views
    1. Add an import:  from other_app.views import Home
    2. Add a URL to urlpatterns:  path('', Home.as_view(), name='home')
Including another URLconf
    1. Import the include() function: from django.urls import include, path
    2. Add a URL to urlpatterns:  path('blog/', include('blog.urls'))
"""
from django.contrib import admin

from django.urls import path, include

from django.conf.urls.static import static

from django.conf import settings

from django.views.generic.base import TemplateView
urlpatterns = [
    path('central-admin/', admin.site.urls), # Admin do Django
    path('', include('core.urls')),
    path('', include('processo.urls')),
    path('contas/', include('django.contrib.auth.urls')), # Para fazer a autenticação de Login/Logout
] + static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)

# Editar a Área Administrativa do Django
admin.site.site_header = 'WebGer23' # Título da tela de Login
admin.site.site_title = 'WebGer23' # Título da aba do navegador da tela de Login
admin.site.index_title = 'Sistema Web Gerenciador de Processos' # Título da página principal após o login