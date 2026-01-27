from django.contrib import admin
from django.contrib.auth.admin import UserAdmin
from django.contrib.auth.models import User, Group
from .models import Requisition, RequisitionItem, AutoUploadConfig, MaterialProcessTypeRule

class RequisitionItemInline(admin.TabularInline):
    model = RequisitionItem
    fields = ('item_name', 'required_quantity', 'material_number', 'confirmed_quantity', 'is_signed_off', 'dispatch_status')
    readonly_fields = ('is_signed_off',) # confirmed_quantity is editable via material_confirmation view
    extra = 0

@admin.register(Requisition)
class RequisitionAdmin(admin.ModelAdmin):
    list_display = ('order_number', 'applicant', 'request_date', 'process_type', 'status', 'created_at', 'remarks')
    list_filter = ('status', 'request_date', 'process_type', 'created_at')
    search_fields = ('order_number', 'applicant__username', 'remarks')
    raw_id_fields = ('applicant',)
    inlines = [RequisitionItemInline]

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

@admin.register(AutoUploadConfig)
class AutoUploadConfigAdmin(admin.ModelAdmin):
    list_display = ('priority', 'get_upload_type_display', 'file_path', 'is_active', 'last_run', 'last_status')
    list_editable = ('is_active', 'file_path') # Removed 'priority' from list_editable
    ordering = ('priority',)

@admin.register(MaterialProcessTypeRule)
class MaterialProcessTypeRuleAdmin(admin.ModelAdmin):
    list_display = ('material_prefix', 'machine_model_name', 'parent_material_desc_keyword', 'process_type_name', 'updated_by', 'updated_at')
    list_filter = ('machine_model_name', 'process_type_name')
    search_fields = ('material_prefix', 'machine_model_name', 'process_type_name', 'parent_material_desc_keyword')
    raw_id_fields = ('updated_by',)
    list_editable = ('parent_material_desc_keyword', 'process_type_name')


from .models import OperationProcessRule

@admin.register(OperationProcessRule)
class OperationProcessRuleAdmin(admin.ModelAdmin):
    list_display = ('operation_description', 'process_type', 'updated_by', 'updated_at')
    list_filter = ('process_type',)
    search_fields = ('operation_description',)
    list_editable = ('process_type',)
    raw_id_fields = ('updated_by',)

