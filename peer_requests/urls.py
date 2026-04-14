from django.urls import path
from . import views

app_name = 'peer_requests'

urlpatterns = [
    path('', views.peer_request_list, name='list'),
    path('history/', views.peer_request_history, name='history'),
    path('create/', views.peer_request_create, name='create'),
    path('accept/<int:pk>/', views.peer_request_accept, name='accept'),
    path('reply/<int:pk>/', views.peer_request_reply, name='reply'),
    path('close/<int:pk>/', views.peer_request_close, name='close'),
    path('save-template/', views.save_peer_request_template, name='save_template'),
    path('delete-template/<int:pk>/', views.delete_peer_request_template, name='delete_template'),
]

