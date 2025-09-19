from django.urls import path
from . import views

urlpatterns = [
    # Master Material List and Import
    path('materials/', views.material_list, name='material_list'),
    path('materials/update_quantities/', views.update_material_quantities, name='update_quantities'),
    path('materials/import/', views.import_material_master, name='import_material_master'),
    path('materials/create_stocktake/', views.create_stocktake_from_selection, name='create_stocktake_from_selection'),

    path('materials/export_differences/', views.export_master_material_differences, name='export_master_material_differences'),

    # Stocktake List and Detail
    path('stocktakes/', views.stocktake_list, name='stocktake_list'),
    path('stocktakes/<int:pk>/', views.stocktake_detail, name='stocktake_detail'),
    path('stocktakes/<int:pk>/update_items/', views.handle_stocktake_actions, name='update_stocktake_items'),
    path('stocktakes/<int:pk>/export_differences/', views.export_stocktake_differences, name='export_stocktake_differences'),

    # Set a default path for the inventory app, perhaps redirect to material_list
    path('', views.material_list, name='inventory_home'),
]