# maintenance_project/urls.py
from django.contrib import admin
from django.urls import path, include

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', include('maintenance_app.urls')),  # Inclut les URLs de l'application
]