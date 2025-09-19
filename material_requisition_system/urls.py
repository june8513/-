from django.contrib import admin
from django.urls import path, include
from requisitions.views import homepage # Import the homepage view
from django.conf import settings
from django.conf.urls.static import static

urlpatterns = [
    path('admin/', admin.site.urls),
    path('accounts/', include('django.contrib.auth.urls')),
    path('requisitions/', include('requisitions.urls')),
    path('inventory/', include('inventory.urls')),
    path('', homepage, name='homepage'), # Map root URL to homepage view
]

if settings.DEBUG:
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
    urlpatterns += static(settings.STATIC_URL, document_root=settings.STATIC_ROOT)
