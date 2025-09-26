from django.urls import path
from . import views

# URL patterns for the redesigned inventory management system
urlpatterns = [
    # Main dashboard
    path('', views.inventory_dashboard, name='inventory_home'),

    # Feature 1: Inventory Update
    path('update/', views.inventory_update_view, name='inventory_update'),
    path('run_update/', views.import_material_master, name='run_inventory_update'), # This is the action URL for the form

    # Feature 2: Stocktake by Location (Placeholders for now)
    path('stocktake/', views.stocktake_location_list, name='stocktake_location_list'),
    path('stocktake/<str:location_name>/', views.stocktake_detail_by_location, name='stocktake_detail_by_location'),
    path('stocktake/update_count/', views.update_counted_quantity, name='update_counted_quantity'),

    # Feature 3: Difference Report (Placeholders for now)
    path('differences/', views.difference_location_list, name='difference_location_list'),
    path('differences/<str:location_name>/', views.difference_detail_by_location, name='difference_detail_by_location'),
    path('differences/<str:location_name>/export/', views.export_differences_excel, name='export_differences_excel'),

    # The old views below are kept for now but are not directly accessible through the new dashboard.
    # They might be removed later.
    path('materials/', views.material_list, name='material_list'),
    path('materials/update_quantities/', views.update_material_quantities, name='update_quantities'),
    path('materials/create_stocktake/', views.create_stocktake_from_selection, name='create_stocktake_from_selection'),
    path('materials/export_differences/', views.export_master_material_differences, name='export_master_material_differences'),
    path('stocktakes/', views.stocktake_list, name='stocktake_list'),
    path('stocktakes/<int:pk>/', views.stocktake_detail, name='stocktake_detail'),
    path('stocktakes/<int:pk>/update_items/', views.handle_stocktake_actions, name='update_stocktake_items'),
    path('stocktakes/<int:pk>/export_differences/', views.export_stocktake_differences, name='export_stocktake_differences'),
]
