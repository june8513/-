from django.urls import path
from . import views
from .views import api_views, assistant_views, simple_views  # Import simple_views
from .views.requisition_management_views import dismiss_requisition_alert

app_name = 'requisitions'

urlpatterns = [
    # Simple interface URLs - 簡易介面路由
    path('simple/applicant/', simple_views.simple_applicant_home, name='simple_applicant_home'),
    path('simple/applicant/create/', simple_views.simple_applicant_create, name='simple_applicant_create'),
    path('simple/applicant/<int:pk>/', simple_views.simple_applicant_detail, name='simple_applicant_detail'),
    path('simple/applicant/<int:pk>/sign-off/', simple_views.simple_applicant_sign_off, name='simple_applicant_sign_off'),
    path('simple/applicant/<int:pk>/update-process-type/', simple_views.simple_applicant_update_process_type, name='simple_applicant_update_process_type'),
    path('simple/applicant/<int:pk>/update-request-date/', simple_views.simple_applicant_update_request_date, name='simple_applicant_update_request_date'),
    path('simple/applicant/<int:pk>/delete/', simple_views.simple_applicant_delete, name='simple_applicant_delete'),
    path('simple/dispatcher/', simple_views.simple_dispatcher_home, name='simple_dispatcher_home'),
    path('simple/dispatcher/<str:category>/', simple_views.simple_dispatcher_category, name='simple_dispatcher_category'),
    path('simple/dispatcher/<str:category>/<int:pk>/', simple_views.simple_dispatcher_detail, name='simple_dispatcher_detail'),

    path('work_orders/', views.work_order_list, name='work_order_list'),
    path('work_orders/<str:order_number>/toggle_archive/', views.toggle_work_order_archive, name='toggle_work_order_archive'),
    path('work_orders/<str:order_number>/requisitions/', views.work_order_requisitions_list, name='work_order_requisitions_list'),
    path('assistant/', assistant_views.assistant_view, name='assistant'),
    path('api/natural_action/', api_views.natural_action_view, name='natural_action'),
    path('api/check_task_status/<str:task_id>/', api_views.check_task_status, name='check_task_status'),
    path('finished_goods_dispatch/', views.finished_goods_dispatch, name='finished_goods_dispatch'),
    path('<int:pk>/images/', views.view_requisition_images, name='view_requisition_images'),
    path('<int:pk>/upload_page/', views.upload_requisition_images_page, name='upload_requisition_images_page'),

    path('batch_dispatch_view/', views.batch_dispatch_view, name='batch_dispatch_view'),
    path('batch_dispatch_action/', views.batch_dispatch_action, name='batch_dispatch_action'),
    path('list/', views.requisition_list, name='requisition_list'),
    path('<int:pk>/detail/', views.requisition_detail, name='requisition_detail'),
    path('archived_list/', views.archived_requisition_list, name='archived_requisition_list'), # Renamed from ''
    path('list/export/excel/', views.export_requisitions_excel, name='export_requisitions_excel'),
    path('archived_list/export/excel/', views.export_archived_requisitions_excel, name='export_archived_requisitions_excel'),
    path('material_list/export/excel/', views.export_work_order_materials_excel, name='export_work_order_materials_excel'),
    path('archived_material_list/export/excel/', views.export_archived_work_order_materials_excel, name='export_archived_work_order_materials_excel'),
    path('list/export/all_pending_materials/excel/', views.export_all_pending_materials_excel, name='export_all_pending_materials_excel'),
    path('create/', views.requisition_create, name='requisition_create'),
    path('get_process_types/', views.get_available_process_types, name='get_available_process_types'),
    

    path('upload_order_model_excel/', views.upload_order_model_excel, name='upload_order_model_excel'),
    path('upload_material_details_excel/', views.upload_material_details_excel, name='upload_material_details_excel'),

    path('<int:pk>/sign_off/', views.requisition_sign_off, name='requisition_sign_off'),
    path('<int:pk>/delete/', views.requisition_delete, name='requisition_delete'), # New URL for deleting requisition
    path('history/', views.requisition_history, name='requisition_history'),
    path('login/', views.user_login, name='login'),
    path('logout/', views.user_logout, name='logout'),
    path('register/', views.user_register, name='register'),
    path('<int:pk>/sign_off_item/<int:item_pk>/', views.sign_off_item, name='sign_off_item'),
    path('update_db/', views.update_process_type_db, name='update_process_type_db'),
    path('upload_inventory/', views.upload_inventory_data, name='upload_inventory_data'),
    path('bulk_upload/', views.bulk_upload, name='bulk_upload'),
    
    # 作業說明投料點規則
    path('classify_operations/', views.classify_operations, name='classify_operations'),
    path('operation_rules/', views.operation_rules_list, name='operation_rules_list'),
    
    # 半成品投料點管理
    path('semi_finished/process_types/', views.semi_finished_process_type_list, name='semi_finished_process_type_list'),

    path('material_list/', views.work_order_material_list, name='work_order_material_list'),
    path('archived_material_list/', views.archived_work_order_material_list, name='archived_work_order_material_list'),
    path('update_work_order_quantities/', views.update_work_order_quantities, name='update_work_order_quantities'),

    path('<int:pk>/generate_dispatch_note/', views.generate_dispatch_note, name='generate_dispatch_note'),
    path('<int:pk>/generate_dispatch_note/excel/', views.generate_dispatch_note, name='generate_dispatch_note_excel'),
    path('<int:pk>/update_dispatch_note/', views.update_dispatch_note, name='update_dispatch_note'),
    path('<int:pk>/update_material_dispatch_status/', views.update_material_dispatch_status, name='update_material_dispatch_status'),
    path('<int:pk>/generate_backorder_note/', views.generate_backorder_note, name='generate_backorder_note'),
    path('<int:pk>/generate_backorder_note/excel/', views.export_backorder_note_excel, name='generate_backorder_note_excel'),
    path('<int:pk>/dismiss_alert/', dismiss_requisition_alert, name='dismiss_requisition_alert'), # New URL
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
    path('work_order_material/<str:material_code>/images/', views.view_work_order_material_images, name='view_work_order_material_images'),
    path('shortage_materials/', views.shortage_materials_list, name='shortage_materials_list'),
    path('shortage_materials/update_dates/', views.update_shortage_arrival_dates, name='update_shortage_arrival_dates'),
    path('update_estimated_arrival_date/', views.update_estimated_arrival_date, name='update_estimated_arrival_date'),
    path('estimated_material_demand/', views.estimated_material_demand, name='estimated_material_demand'),
    path('dispatch_note/<int:pk>/upload_image/', views.upload_dispatch_note_image, name='upload_dispatch_note_image'),
    path('material_completeness/', views.material_completeness, name='material_completeness'),
]
