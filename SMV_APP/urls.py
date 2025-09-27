"""
URL configuration for SMV_APP project.

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/5.1/topics/http/urls/
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
from django.urls import path
from .view import *

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', acceder,name="index"),
    path('descargar-datos-financieros/', descargar_datos_financieros, name='descargar_datos_financieros'),
    path('delete-file/', delete_file, name='delete_file'),
    path('verificar-archivos/', verificar_archivos, name='verificar_archivos'),
      path('preview-excel/', preview_excel, name='preview_excel'),
]
