from django.contrib import admin
from .models import PeerRequest

@admin.register(PeerRequest)
class PeerRequestAdmin(admin.ModelAdmin):
    list_display = ('id', 'applicant', 'recipient', 'request_date', 'status', 'created_at')
    list_filter = ('status', 'request_date')
    search_fields = ('applicant__username', 'recipient__username', 'description')
    readonly_fields = ('created_at', 'updated_at', 'closed_at')
