from django.urls import path
from . import views

app_name = 'core'

urlpatterns = [
    path('', views.homepage, name='homepage'),
    path('supervisor/process/<str:process_name>/', views.supervisor_process_detail, name='supervisor_process_detail'),
]
