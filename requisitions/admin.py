from django.contrib import admin
from django.contrib.auth.admin import UserAdmin
from django.contrib.auth.models import User, Group
from .models import Requisition, MaterialListVersion, RequisitionItem

class RequisitionItemInline(admin.TabularInline):
    model = RequisitionItem
    fields = ('material_name', 'quantity', 'material_number', 'confirmed_quantity', 'is_signed_off')
    readonly_fields = ('is_signed_off',) # confirmed_quantity is editable via material_confirmation view
    extra = 0

@admin.register(MaterialListVersion)
class MaterialListVersionAdmin(admin.ModelAdmin):
    list_display = ('requisition', 'uploaded_at', 'uploaded_by', 'is_active_version')
    list_filter = ('uploaded_at', 'uploaded_by')
    search_fields = ('requisition__order_number', 'uploaded_by__username')
    inlines = [RequisitionItemInline]
    raw_id_fields = ('requisition', 'uploaded_by')

    def is_active_version(self, obj):
        return obj == obj.requisition.current_material_list_version
    is_active_version.boolean = True
    is_active_version.short_description = '當前版本'

@admin.register(Requisition)
class RequisitionAdmin(admin.ModelAdmin):
    list_display = ('order_number', 'applicant', 'request_date', 'process_type', 'status', 'created_at', 'remarks')
    list_filter = ('status', 'request_date', 'process_type', 'created_at')
    search_fields = ('order_number', 'applicant__username', 'remarks')
    raw_id_fields = ('applicant',)

    fieldsets = (
        (None, {
            'fields': ('order_number', 'applicant', 'request_date', 'process_type', 'status', 'remarks')
        }),
    )

    def get_queryset(self, request):
        qs = super().get_queryset(request)
        return qs

    def save_model(self, request, obj, form, change):
        if not obj.pk and not obj.applicant:
            obj.applicant = request.user
        super().save_model(request, obj, form, change)

# Unregister the default User admin
admin.site.unregister(User)

@admin.register(User)
class CustomUserAdmin(UserAdmin):
    # Reconstruct fieldsets to avoid duplication
    fieldsets = (
        (None, {'fields': ('username', 'password')}),
        ('個人資訊', {'fields': ('first_name', 'last_name', 'email')}),
        ('權限', {'fields': ('is_active', 'is_staff', 'is_superuser', 'user_permissions')}),
        ('重要日期', {'fields': ('last_login', 'date_joined')}),
        ('角色', {'fields': ('groups',)}), # Add groups here
    )

    # Customize list_display to show group membership
    list_display = UserAdmin.list_display + ('get_groups',)

    def get_groups(self, obj):
        return ", ".join([g.name for g in obj.groups.all()])
    get_groups.short_description = '所屬角色'
