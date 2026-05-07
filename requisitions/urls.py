from django.urls import path
from .views import (
    auth_views,
    dispatch_views,
    image_views,
    material_data_views,
    material_demand_views,
    material_export_views,
    material_upload_views,
    requisition_management_views,
    api_views,
    assistant_views,
    simple_views,
    work_order_views,
    operation_rule_views,
    semi_finished_views,
    material_completeness_views,
)

app_name = 'requisitions'

urlpatterns = [
    # Simple interface URLs - 簡易介面路由
    # Fast Dispatch (快速撥料)
    path('simple/dispatcher/fast-dispatch/', simple_views.simple_dispatcher_fast_dispatch, name='simple_dispatcher_fast_dispatch'),
    path('simple/dispatcher/fast-dispatch/execute/', simple_views.simple_dispatcher_fast_dispatch_execute, name='simple_dispatcher_fast_dispatch_execute'),

    # Shortage Inquiry (缺料查詢)
    path('simple/shortage-inquiry/', simple_views.shortage_inquiry, name='shortage_inquiry'),
    path('simple/shortage-inquiry/export/', simple_views.shortage_inquiry_export, name='shortage_inquiry_export'),
    path('simple/shortage-inquiry/sync-mps/', simple_views.sync_mps_order_info, name='sync_mps_order_info'),


    path('simple/applicant/', simple_views.simple_applicant_home, name='simple_applicant_home'),
    path('simple/applicant/create/', simple_views.simple_applicant_create, name='simple_applicant_create'),
    path('simple/applicant/<int:pk>/', simple_views.simple_applicant_detail, name='simple_applicant_detail'),
    path('simple/applicant/<int:pk>/sign-off/', simple_views.simple_applicant_sign_off, name='simple_applicant_sign_off'),
    path('simple/applicant/item/<int:item_id>/report-issue/', simple_views.report_item_issue, name='report_item_issue'),
    path('simple/applicant/<int:pk>/update-process-type/', simple_views.simple_applicant_update_process_type, name='simple_applicant_update_process_type'),
    path('simple/applicant/<int:pk>/update-request-date/', simple_views.simple_applicant_update_request_date, name='simple_applicant_update_request_date'),
    path('simple/applicant/<int:pk>/delete/', simple_views.simple_applicant_delete, name='simple_applicant_delete'),
    path('simple/dispatcher/', simple_views.simple_dispatcher_home, name='simple_dispatcher_home'),
    path('simple/dispatcher/<str:category>/', simple_views.simple_dispatcher_category, name='simple_dispatcher_category'),
    path('simple/dispatcher/<str:category>/shortage/', simple_views.simple_dispatcher_shortage, name='simple_dispatcher_shortage'),
    path('simple/dispatcher/<str:category>/merge/', simple_views.simple_dispatcher_merge, name='simple_dispatcher_merge'),
    path('simple/dispatcher/<str:category>/<int:pk>/', simple_views.simple_dispatcher_detail, name='simple_dispatcher_detail'),
    path('simple/dispatcher/item/<int:item_id>/resolve-issue/', simple_views.resolve_item_issue, name='resolve_item_issue'),
    path('simple/dispatcher/item/<int:item_id>/dispatch/', simple_views.simple_dispatch_item_ajax, name='simple_dispatch_item_ajax'),
    path('simple/dispatcher/announcement/update/', simple_views.update_announcement, name='update_announcement'),

    # Excel export for simple interface
    path('simple/applicant/export/excel/', simple_views.export_simple_applicant_requisitions_excel, name='export_simple_applicant_requisition_excel'),
    path('simple/dispatcher/<str:category>/export/excel/', simple_views.export_simple_dispatcher_requisitions_excel, name='export_simple_dispatcher_requisition_excel'),
    path('simple/requisition/<int:pk>/export/excel/', simple_views.export_single_requisition_excel, name='export_single_requisition_excel'),
    path('simple/requisition/<int:pk>/change-alert/', simple_views.simple_requisition_change_detail, name='simple_requisition_change_detail'),
    path('simple/requisition-item/<int:item_pk>/dismiss-alert/', simple_views.dismiss_requisition_item_alert, name='dismiss_requisition_item_alert'),

    path('work_orders/', work_order_views.work_order_list, name='work_order_list'),
    path('work_orders/<str:order_number>/toggle_archive/', work_order_views.toggle_work_order_archive, name='toggle_work_order_archive'),
    path('work_orders/<str:order_number>/requisitions/', work_order_views.work_order_requisitions_list, name='work_order_requisitions_list'),
    path('assistant/', assistant_views.assistant_view, name='assistant'),
    path('api/natural_action/', api_views.natural_action_view, name='natural_action'),
    path('api/check_task_status/<str:task_id>/', api_views.check_task_status, name='check_task_status'),
    path('api/save_ai_correction/', api_views.save_ai_correction, name='save_ai_correction'),
    path('api/shortage_materials/', api_views.shortage_materials_api, name='shortage_materials_api'),
    path('api/requisition_items/shortages/', api_views.requisition_items_shortages_api, name='requisition_items_shortages_api'),
    path('finished_goods_dispatch/', dispatch_views.finished_goods_dispatch, name='finished_goods_dispatch'),
    path('<int:pk>/images/', image_views.view_requisition_images, name='view_requisition_images'),
    path('<int:pk>/upload_page/', image_views.upload_requisition_images_page, name='upload_requisition_images_page'),

    path('batch_dispatch_view/', dispatch_views.batch_dispatch_view, name='batch_dispatch_view'),
    path('batch_dispatch_action/', dispatch_views.batch_dispatch_action, name='batch_dispatch_action'),
    path('list/', requisition_management_views.requisition_list, name='requisition_list'),
    path('<int:pk>/detail/', requisition_management_views.requisition_detail, name='requisition_detail'),
    path('archived_list/', requisition_management_views.archived_requisition_list, name='archived_requisition_list'), # Renamed from ''
    path('list/export/excel/', material_export_views.export_requisitions_excel, name='export_requisitions_excel'),
    path('archived_list/export/excel/', material_export_views.export_archived_requisitions_excel, name='export_archived_requisitions_excel'),
    path('material_list/export/excel/', material_export_views.export_work_order_materials_excel, name='export_work_order_materials_excel'),
    path('archived_material_list/export/excel/', material_export_views.export_archived_work_order_materials_excel, name='export_archived_work_order_materials_excel'),
    path('list/export/all_pending_materials/excel/', material_export_views.export_all_pending_materials_excel, name='export_all_pending_materials_excel'),
    path('create/', requisition_management_views.requisition_create, name='requisition_create'),
    path('get_process_types/', requisition_management_views.get_available_process_types, name='get_available_process_types'),
    

    path('upload_order_model_excel/', material_upload_views.upload_order_model_excel, name='upload_order_model_excel'),
    path('upload_material_details_excel/', material_upload_views.upload_material_details_excel, name='upload_material_details_excel'),

    path('<int:pk>/sign_off/', requisition_management_views.requisition_sign_off, name='requisition_sign_off'),
    path('<int:pk>/delete/', requisition_management_views.requisition_delete, name='requisition_delete'), # New URL for deleting requisition
    path('history/', requisition_management_views.requisition_history, name='requisition_history'),
    path('login/', auth_views.user_login, name='login'),
    path('logout/', auth_views.user_logout, name='logout'),
    path('register/', auth_views.user_register, name='register'),
    path('<int:pk>/sign_off_item/<int:item_pk>/', requisition_management_views.sign_off_item, name='sign_off_item'),
    path('update_db/', material_upload_views.update_process_type_db, name='update_process_type_db'),
    path('upload_inventory/', material_upload_views.upload_inventory_data, name='upload_inventory_data'),
    path('bulk_upload/', material_upload_views.bulk_upload, name='bulk_upload'),
    
    # 作業說明投料點規則
    path('classify_operations/', operation_rule_views.classify_operations, name='classify_operations'),
    path('operation_rules/', operation_rule_views.operation_rules_list, name='operation_rules_list'),
    
    # 半成品投料點管理
    path('semi_finished/process_types/', semi_finished_views.semi_finished_process_type_list, name='semi_finished_process_type_list'),

    path('material_list/', material_data_views.work_order_material_list, name='work_order_material_list'),
    path('archived_material_list/', material_data_views.archived_work_order_material_list, name='archived_work_order_material_list'),
    path('update_work_order_quantities/', material_data_views.update_work_order_quantities, name='update_work_order_quantities'),
    path('sync_storage_bins/', material_data_views.sync_storage_bins, name='sync_storage_bins'),
    path('get_model_process_type_history/', material_data_views.get_model_process_type_history, name='get_model_process_type_history'),

    path('<int:pk>/generate_dispatch_note/', material_export_views.generate_dispatch_note, name='generate_dispatch_note'),
    path('<int:pk>/generate_dispatch_note/excel/', material_export_views.generate_dispatch_note, name='generate_dispatch_note_excel'),
    path('<int:pk>/update_dispatch_note/', dispatch_views.update_dispatch_note, name='update_dispatch_note'),
    path('<int:pk>/update_material_dispatch_status/', dispatch_views.update_material_dispatch_status, name='update_material_dispatch_status'),
    path('<int:pk>/generate_backorder_note/', dispatch_views.generate_backorder_note, name='generate_backorder_note'),
    path('<int:pk>/generate_backorder_note/excel/', material_export_views.export_backorder_note_excel, name='generate_backorder_note_excel'),
    path('<int:pk>/dismiss_alert/', requisition_management_views.dismiss_requisition_alert, name='dismiss_requisition_alert'), # New URL
    path('database/', material_data_views.view_database, name='view_database'),
    path('inventory_database/', material_data_views.view_inventory_database, name='inventory_database'),
    path('clear_database/', material_data_views.clear_work_order_material_database, name='clear_work_order_material_database'),
    path('process_type_database/', material_data_views.view_process_type_database, name='view_process_type_database'),
    path('update_material_process_type/<int:material_id>/', material_data_views.update_material_process_type, name='update_material_process_type'),
    path('get_process_types_for_model/', material_data_views.get_process_types_for_model, name='get_process_types_for_model'),
    path('api/get_process_types_for_order/', material_data_views.get_process_types_for_order, name='get_process_types_for_order'),
    path('process_types_management/', material_data_views.process_types_management, name='process_types_management'),
    path('<int:pk>/details_json/', requisition_management_views.get_requisition_details_json, name='get_requisition_details_json'), # New URL for JSON details
    path('<int:pk>/images_json/', requisition_management_views.get_requisition_images_json, name='get_requisition_images_json'), # New URL for images JSON
    path('<int:pk>/supplement/', dispatch_views.supplement_material, name='supplement_material'),
    path('<int:pk>/upload_images/', image_views.upload_requisition_images, name='upload_requisition_images'),
    path('work_order_material/<int:pk>/upload_images/', image_views.upload_work_order_material_images, name='upload_work_order_material_images'),
    path('work_order_material/<str:material_code>/images/', image_views.view_work_order_material_images, name='view_work_order_material_images'),
    path('shortage_materials/', material_demand_views.shortage_materials_list, name='shortage_materials_list'),
    path('shortage_materials/update_dates/', material_demand_views.update_shortage_arrival_dates, name='update_shortage_arrival_dates'),
    path('update_estimated_arrival_date/', material_demand_views.update_estimated_arrival_date, name='update_estimated_arrival_date'),
    path('estimated_material_demand/', material_demand_views.estimated_material_demand, name='estimated_material_demand'),
    path('dispatch_note/<int:pk>/upload_image/', image_views.upload_dispatch_note_image, name='upload_dispatch_note_image'),
    path('material_completeness/', material_completeness_views.material_completeness, name='material_completeness'),
]
