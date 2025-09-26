from django import forms
from django.forms import modelformset_factory, ClearableFileInput
from .models import Requisition, RequisitionItem, ProcessType, MachineModel

from django.core.exceptions import ValidationError # Import ValidationError

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



class RequisitionImageUploadForm(forms.Form):
    pass

class StagedBulkUploadMaterialsForm(forms.Form):
    file = forms.FileField(label='選擇分階段批量物料 Excel 檔案')

class WorkOrderMaterialImageUploadForm(forms.Form):
    process_type = forms.ModelChoiceField(
        queryset=ProcessType.objects.all(),
        label="投料點",
        required=False,
        empty_label="所有投料點"
    )
    images = forms.FileField(
        label='選擇圖片',
        required=False,
        widget=forms.ClearableFileInput() # Remove multiple=True
    )
