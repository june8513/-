from django.urls import path
from . import views

urlpatterns = [
    path('', views.homepage, name='homepage'), # New homepage
    path('list/', views.requisition_list, name='requisition_list'), # Renamed from ''
    path('list/export/excel/', views.export_requisitions_excel, name='export_requisitions_excel'),
    path('material_list/export/excel/', views.export_work_order_materials_excel, name='export_work_order_materials_excel'),
    path('list/export/all_pending_materials/excel/', views.export_all_pending_materials_excel, name='export_all_pending_materials_excel'),
    path('create/', views.requisition_create, name='requisition_create'),
    path('get_process_types/', views.get_available_process_types, name='get_available_process_types'),
    
    path('<int:pk>/upload_materials/', views.upload_materials, name='upload_materials'),
    path('upload_order_model_excel/', views.upload_order_model_excel, name='upload_order_model_excel'),
    path('upload_material_details_excel/', views.upload_material_details_excel, name='upload_material_details_excel'),
    path('<int:pk>/material_confirmation/', views.material_confirmation, name='material_confirmation'),
    path('<int:pk>/material_confirmation/export/excel/', views.export_material_confirmation_excel, name='export_material_confirmation_excel'), # New URL for material handler confirmation
<<<<<<< HEAD

    path('<int:pk>/delete/', views.requisition_delete, name='requisition_delete'), # New URL for deleting requisition
    path('history/', views.requisition_history, name='requisition_history'),

=======
    path('<int:pk>/sign_off/<int:version_pk>/', views.requisition_sign_off, name='requisition_sign_off_version'),
    path('<int:pk>/sign_off/', views.requisition_sign_off, name='requisition_sign_off'),
    path('<int:pk>/delete/', views.requisition_delete, name='requisition_delete'), # New URL for deleting requisition
    path('history/', views.requisition_history, name='requisition_history'),
    path('<int:pk>/detail/', views.requisition_detail, name='requisition_detail'), # New URL for requisition detail
>>>>>>> 4ac9e3d0ff5915a8953899870be6616b6f0653c9
    path('login/', views.user_login, name='login'),
    path('logout/', views.user_logout, name='logout'),
    path('<int:pk>/activate_version/<int:version_pk>/', views.activate_material_version, name='activate_material_version'),
    path('<int:pk>/sign_off_item/<int:version_pk>/<int:item_pk>/', views.sign_off_item, name='sign_off_item'),
    path('update_db/', views.update_process_type_db, name='update_process_type_db'),
    path('upload_inventory/', views.upload_inventory_data, name='upload_inventory_data'),
    path('upload_storage_bin/', views.upload_storage_bin_data, name='upload_storage_bin_data'),
    path('material_list/', views.work_order_material_list, name='work_order_material_list'),
    path('update_work_order_quantities/', views.update_work_order_quantities, name='update_work_order_quantities'),
    path('import_materials/', views.import_materials_to_requisition, name='import_materials_to_requisition'),
    path('<int:pk>/generate_dispatch_note/', views.generate_dispatch_note, name='generate_dispatch_note'),
    path('<int:pk>/generate_dispatch_note/excel/', views.generate_dispatch_note, name='generate_dispatch_note_excel'),
    path('<int:pk>/update_dispatch_note/', views.update_dispatch_note, name='update_dispatch_note'),
    path('<int:pk>/update_material_dispatch_status/', views.update_material_dispatch_status, name='update_material_dispatch_status'),
    path('<int:pk>/generate_backorder_note/', views.generate_backorder_note, name='generate_backorder_note'),
    path('<int:pk>/generate_backorder_note/excel/', views.export_backorder_note_excel, name='generate_backorder_note_excel'),
    path('database/', views.view_database, name='view_database'),
    path('inventory_database/', views.view_inventory_database, name='inventory_database'),
    path('clear_database/', views.clear_work_order_material_database, name='clear_work_order_material_database'),
    path('process_type_database/', views.view_process_type_database, name='view_process_type_database'),
    path('update_material_process_type/<int:material_id>/', views.update_material_process_type, name='update_material_process_type'),
    path('get_process_types_for_model/', views.get_process_types_for_model, name='get_process_types_for_model'),
    path('process_types_management/', views.process_types_management, name='process_types_management'),
    path('<int:pk>/details_json/', views.get_requisition_details_json, name='get_requisition_details_json'), # New URL for JSON details
    path('<int:pk>/images_json/', views.get_requisition_images_json, name='get_requisition_images_json'), # New URL for images JSON
    path('<int:pk>/supplement/', views.supplement_material, name='supplement_material'),
    path('<int:pk>/upload_images/', views.upload_requisition_images, name='upload_requisition_images'),
    path('work_order_material/<int:pk>/upload_images/', views.upload_work_order_material_images, name='upload_work_order_material_images'),
    path('shortage_materials/', views.shortage_materials_list, name='shortage_materials_list'),
]
