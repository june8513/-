from django.urls import path
from . import views

app_name = 'interactive_picking'

urlpatterns = [
    path('', views.index, name='index'),
    path('wizard/<int:task_id>/', views.picking_wizard, name='wizard'),
    path('action/<int:item_id>/', views.process_picking_action, name='process_action'),
    path('undo/<int:task_id>/', views.undo_last_action, name='undo_action'),
    path('setup-locations/', views.upload_warehouse_locations, name='upload_locations'),
]
