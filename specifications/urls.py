from django.urls import path
from . import views

urlpatterns = [
    path('', views.material_spec_list, name='material_spec_list'),
    path('import/', views.import_material_specs, name='import_material_specs'),
    path('<int:material_id>/edit/', views.material_spec_edit, name='material_spec_edit'),
]
