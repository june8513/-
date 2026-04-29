from django.contrib import admin
from django.contrib.auth.admin import UserAdmin
from django.contrib.auth.models import User, Group
from .models import (
    Requisition, RequisitionItem, AutoUploadConfig, MaterialProcessTypeRule, 
    UserProfile, RequisitionShareGroup, AIUserCorrection, Announcement,
    WorkOrder, WorkOrderMaterial, MachineModel, ProcessType, SemiFinishedProcessType,
    WorkOrderMaterialTransaction, WorkOrderMaterialProcessTypeLog
)

class RequisitionItemInline(admin.TabularInline):
    model = RequisitionItem
    fields = ('item_name', 'required_quantity', 'material_number', 'confirmed_quantity', 'is_signed_off', 'dispatch_status')
    # Removed is_signed_off from readonly_fields to allow manual adjustment in back-end
    readonly_fields = () 
    extra = 0

class UserProfileInline(admin.StackedInline):
    model = UserProfile
    can_delete = False
    verbose_name_plural = '使用者設定檔'
    fields = ('requested_role', 'can_publish_announcements')

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

    inlines = (UserProfileInline, )

    # Customize list_display to show group membership and requested role
    list_display = ('username', 'email', 'first_name', 'last_name', 'is_active', 'get_groups', 'get_requested_role')
    list_filter = UserAdmin.list_filter + ('groups', 'profile__requested_role')
    actions = ['approve_users']

    def get_groups(self, obj):
        return ", ".join([g.name for g in obj.groups.all()])
    get_groups.short_description = '所屬角色'

    def get_requested_role(self, obj):
        if hasattr(obj, 'profile') and obj.profile.requested_role:
            return obj.profile.requested_role
        return '-'
    get_requested_role.short_description = '申請角色 (待審核)'

    def approve_users(self, request, queryset):
        success_count = 0
        fail_count = 0
        
        for user in queryset:
            if hasattr(user, 'profile') and user.profile.requested_role:
                role_name = user.profile.requested_role
                try:
                    group = Group.objects.get(name=role_name)
                    user.groups.add(group)
                    user.is_active = True
                    user.save()
                    success_count += 1
                except Group.DoesNotExist:
                    self.message_user(request, f"使用者 {user.username} 申請的角色「{role_name}」在系統中不存在。", level='ERROR')
                    fail_count += 1
            else:
                # 如果沒有申請角色，但被選中審核，可以直接啟用嗎？
                # 保守起見，僅啟用，不分配角色，並提示
                user.is_active = True
                user.save()
                self.message_user(request, f"使用者 {user.username} 沒有申請角色，已僅將其啟用。", level='WARNING')
                success_count += 1
        
        self.message_user(request, f"已成功審核並啟用 {success_count} 位使用者。")
    approve_users.short_description = "審核通過並分配角色"

@admin.register(AutoUploadConfig)
class AutoUploadConfigAdmin(admin.ModelAdmin):
    list_display = ('priority', 'get_upload_type_display', 'file_path', 'is_active', 'last_run', 'last_status')
    list_editable = ('is_active', 'file_path') 
    ordering = ('priority',)
    actions = ['run_upload_now']

    def run_upload_now(self, request, queryset):
        import io
        from django.core.management import call_command
        from django.utils import timezone
        
        success_count = 0
        
        for config in queryset:
            if not config.file_path:
                self.message_user(request, f"設定 {config} 沒有檔案路徑，跳過。", level='WARNING')
                continue
                
            command_map = {
                'inventory': 'auto_upload_inventory',
                'order_model': 'auto_upload_order_models',
                'material_details': 'auto_upload_material_details',
                'semi_finished': 'auto_upload_semi_finished',
                'semi_finished_model_db': 'auto_upload_semi_finished_model_db',
                'shipping_customer': 'auto_upload_shipping_customer',
                'supplier_data': 'auto_upload_supplier_data',
            }
            
            cmd_name = command_map.get(config.upload_type)
            if not cmd_name:
                self.message_user(request, f"未知的上傳類型: {config.upload_type}", level='ERROR')
                continue
                
            try:
                out = io.StringIO()
                call_command(cmd_name, path=config.file_path, stdout=out)
                
                output_msg = out.getvalue()
                self.message_user(request, f"[{config.get_upload_type_display()}] 執行成功：{output_msg}")
                success_count += 1
                
                # 更新狀態
                config.last_run = timezone.now()
                config.last_status = "Manual Run: Success"
                config.save()
                    
            except Exception as e:
                self.message_user(request, f"執行 {config.get_upload_type_display()} 時發生錯誤: {e}", level='ERROR')
                config.last_status = f"Manual Run Error: {e}"
                config.save()
                
    run_upload_now.short_description = "立即執行自動上傳 (Run Now)"

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


from django import forms
from django.contrib.admin.widgets import FilteredSelectMultiple

class RequisitionShareGroupForm(forms.ModelForm):
    members = forms.ModelMultipleChoiceField(
        queryset=User.objects.all(),
        required=False,
        widget=FilteredSelectMultiple(verbose_name='成員', is_stacked=False),
        label="群組成員"
    )

    class Meta:
        model = RequisitionShareGroup
        fields = '__all__'

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        if 'members' in self.fields:
            self.fields['members'].label_from_instance = lambda obj: f"{obj.get_full_name()} ({obj.username})" if obj.get_full_name() else obj.username

@admin.register(RequisitionShareGroup)
class RequisitionShareGroupAdmin(admin.ModelAdmin):
    form = RequisitionShareGroupForm
    list_display = ('name', 'get_members_count', 'created_at')
    search_fields = ('name', 'members__username', 'members__first_name', 'members__last_name')
    # filter_horizontal = ('members',)
    
    def get_members_count(self, obj):
        return obj.members.count()
    get_members_count.short_description = '成員數量'

@admin.register(AIUserCorrection)
class AIUserCorrectionAdmin(admin.ModelAdmin):
    list_display = ('user', 'query_text_short', 'is_active', 'created_at')
    list_filter = ('is_active', 'created_at', 'user')
    search_fields = ('query_text', 'correction_text', 'incorrect_response')
    list_editable = ('is_active',)

    def query_text_short(self, obj):
        return obj.query_text[:50]
    query_text_short.short_description = '問題內容'

@admin.register(Announcement)
class AnnouncementAdmin(admin.ModelAdmin):
    list_display = ('content_short', 'is_active', 'is_system_generated', 'created_by', 'created_at', 'expires_at')
    list_filter = ('is_active', 'is_system_generated', 'created_at', 'expires_at')
    search_fields = ('content',)
    list_editable = ('is_active',)

    def content_short(self, obj):
        return obj.content[:50]
    content_short.short_description = '公告內容'
@admin.register(MachineModel)
class MachineModelAdmin(admin.ModelAdmin):
    list_display = ('name',)
    search_fields = ('name',)

@admin.register(ProcessType)
class ProcessTypeAdmin(admin.ModelAdmin):
    list_display = ('name', 'machine_model', 'parent', 'is_kit')
    list_filter = ('machine_model', 'is_kit')
    search_fields = ('name',)

@admin.register(SemiFinishedProcessType)
class SemiFinishedProcessTypeAdmin(admin.ModelAdmin):
    list_display = ('name', 'color', 'order', 'is_active')
    list_editable = ('order', 'is_active', 'color')
    search_fields = ('name',)

@admin.register(WorkOrder)
class WorkOrderAdmin(admin.ModelAdmin):
    list_display = ('order_number', 'customer_name', 'shipping_date', 'is_archived', 'created_at')
    list_filter = ('is_archived', 'shipping_date')
    search_fields = ('order_number', 'customer_name')
    list_editable = ('is_archived',)

@admin.register(WorkOrderMaterial)
class WorkOrderMaterialAdmin(admin.ModelAdmin):
    list_display = ('order_number', 'material_number', 'item_name', 'required_quantity', 'confirmed_quantity', 'is_signed_off', 'process_type', 'is_active')
    list_filter = ('is_active', 'material_type', 'process_type')
    search_fields = ('order_number', 'material_number', 'item_name')
    list_editable = ('is_signed_off', 'is_active', 'confirmed_quantity')
    raw_id_fields = ('machine_model', 'process_type')

@admin.register(WorkOrderMaterialTransaction)
class WorkOrderMaterialTransactionAdmin(admin.ModelAdmin):
    list_display = ('timestamp', 'work_order_material', 'user', 'transaction_type', 'quantity_change', 'new_confirmed_quantity')
    list_filter = ('transaction_type', 'timestamp', 'user')
    search_fields = ('work_order_material__material_number', 'work_order_material__order_number', 'notes')

@admin.register(WorkOrderMaterialProcessTypeLog)
class WorkOrderMaterialProcessTypeLogAdmin(admin.ModelAdmin):
    list_display = ('timestamp', 'work_order_material', 'user', 'old_process_type', 'new_process_type')
    list_filter = ('timestamp', 'user')
    search_fields = ('work_order_material__material_number', 'work_order_material__order_number')
