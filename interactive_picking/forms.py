from django import forms

class UploadWarehouseLocationForm(forms.Form):
    excel_file = forms.FileField(
        label='選擇包含儲位座標的 Excel 檔案',
        widget=forms.ClearableFileInput(attrs={'class': 'form-control', 'accept': '.xlsx, .xls'})
    )
