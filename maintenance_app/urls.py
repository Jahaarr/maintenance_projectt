# maintenance_app/urls.py
from django.urls import path
from . import views

urlpatterns = [
    path('equipments/', views.equipment_list, name='equipment_list'),
    path('equipments/<int:equipement_id>/', views.equipment_detail, name='equipment_detail'),
    path('update_sous_ensemble/<int:sous_ensemble_id>/', views.update_sous_ensemble, name='update_sous_ensemble'),
    path('export_excel/', views.export_excel, name='export_excel'),
    path('dashboard/', views.dashboard, name='dashboard'),
]