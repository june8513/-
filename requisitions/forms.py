from django import forms
from django.forms import modelformset_factory, ClearableFileInput
from .models import Requisition, RequisitionItem, ProcessType, MachineModel, RequisitionImage, WorkOrderMaterialImage

# Custom widget for multiple file uploads
class MultipleFileInput(forms.FileInput):
    def render(self, name, value, attrs=None, renderer=None):
        final_attrs = self.build_attrs(self.attrs, attrs)
        final_attrs['multiple'] = True
        return super().render(name, value, final_attrs, renderer)

class RequisitionImageForm(forms.Form):
    image = forms.ImageField(
        label='上傳圖片 (可多選)',
        widget=MultipleFileInput()
    )

class RequisitionForm(forms.ModelForm):
    order_number = forms.CharField(
        max_length=100,
        required=True,
        label="訂單單號",
        widget=forms.TextInput(attrs={'placeholder': '請輸入訂單單號'})
    )
    process_type = forms.ChoiceField(
        choices=[], # Empty choices initially, will be populated by JS or constructor
        required=True,
        label="需求流程",
        widget=forms.Select(attrs={'id': 'id_process_type', 'required': 'required'}),
        error_messages={'invalid_choice': '無法申請：所選的需求流程無效或已歸檔。'} # Custom error message
    )

    def __init__(self, *args, **kwargs):
        process_type_choices = kwargs.pop('process_type_choices', None)
        super().__init__(*args, **kwargs)
        if process_type_choices:
            self.fields['process_type'].choices = process_type_choices

    def clean_process_type(self):
        process_type_id = self.cleaned_data['process_type']
        # Check if the submitted process_type_id is actually in the available choices
        # This handles the case where a user might try to submit an invalid choice
        # or a choice that was valid but became invalid (e.g., material archived)
        available_ids = [str(choice[0]) for choice in self.fields['process_type'].choices]
        if process_type_id not in available_ids:
            raise ValidationError("無法申請：所選的需求流程無效或已歸檔。")
        return process_type_id

    class Meta:
        model = Requisition
        fields = ['order_number', 'request_date', 'process_type', 'remarks']
        widgets = {
            'request_date': forms.DateInput(attrs={'type': 'date'}),
            # Remove process_type from here as it's defined explicitly above
        }

class UploadFileForm(forms.Form):
    file = forms.FileField(label='選擇 Excel 檔案')

class OrderModelUploadForm(forms.Form):
    file = forms.FileField(label='選擇訂單與機型 Excel 檔案')

class MaterialDetailsUploadForm(forms.Form):
    file = forms.FileField(label='選擇物料明細 Excel 檔案')
    required_quantity_col = forms.CharField(label='需求數量欄位名稱', initial='需求數量')
    demand_date_col = forms.CharField(label='需求日期欄位名稱', initial='需求日期', required=False)

class UpdateProcessTypeDBForm(forms.Form):
    file = forms.FileField(label='選擇新的投料點資料庫 Excel 檔案 (output.xlsx)')

class UploadInventoryFileForm(forms.Form):
    file = forms.FileField(label='選擇庫存 Excel 檔案')



# Formset for Material Handler's confirmation
class RequisitionItemMaterialConfirmationForm(forms.ModelForm):
    class Meta:
        model = RequisitionItem
        fields = ('confirmed_quantity',)
        widgets = {
            'confirmed_quantity': forms.NumberInput(attrs={'step': '0.01'}), # Allow decimal input
        }

RequisitionItemMaterialConfirmationFormSet = modelformset_factory(
    RequisitionItem,
    form=RequisitionItemMaterialConfirmationForm,
    extra=0, # Do not add extra empty forms
    can_delete=False,
)

# Formset for Applicant's final sign-off
RequisitionItemSignOffFormSet = modelformset_factory(
    RequisitionItem,
    fields=('is_signed_off',),
    extra=0, # Do not add extra empty forms
    can_delete=False,
)

class ProcessTypeForm(forms.ModelForm):
    class Meta:
        model = ProcessType
        fields = ['name', 'machine_model']
        labels = {
            'name': '投料點名稱',
            'machine_model': '所屬機型',
        }

class StagedBulkUploadMaterialsForm(forms.Form):
    file = forms.FileField(label='選擇分階段批量物料 Excel 檔案')

class WorkOrderMaterialImageUploadForm(forms.ModelForm):
    class Meta:
        model = WorkOrderMaterialImage
        fields = ['image']


class BulkUploadForm(forms.Form):
    """一鍵更新：同時上傳庫存、訂單機型、物料明細"""
    inventory_file = forms.FileField(
        label='庫存資料 (零件庫存.xlsx)',
        required=False,
        widget=forms.FileInput(attrs={'accept': '.xlsx,.xls'})
    )
    order_model_file = forms.FileField(
        label='訂單機型 (成品入庫TECO狀態.xlsx)',
        required=False,
        widget=forms.FileInput(attrs={'accept': '.xlsx,.xls'})
    )
    material_details_file = forms.FileField(
        label='成品物料明細 (成品撥料.xlsx)',
        required=False,
        widget=forms.FileInput(attrs={'accept': '.xlsx,.xls'})
    )
    semi_finished_file = forms.FileField(
        label='半成品物料明細 (semi_finished.xlsx)',
        required=False,
        widget=forms.FileInput(attrs={'accept': '.xlsx,.xls'})
    )
    semi_finished_model_file = forms.FileField(
        label='半成品機型資料庫對照表 (訂單-機型)',
        required=False,
        widget=forms.FileInput(attrs={'accept': '.xlsx,.xls'})
    )
    required_quantity_col = forms.CharField(
        label='需求數量欄位名稱',
        initial='需求數量',
        required=False
    )

