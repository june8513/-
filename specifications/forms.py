from django import forms
from .models import MaterialSpecification
from inventory.models import MaterialImage, Material

class MaterialSpecificationForm(forms.ModelForm):
    # Add the 'bin' field from the Material model
    bin = forms.CharField(max_length=100, required=False, label="儲格")

    class Meta:
        model = MaterialSpecification
        fields = ['size', 'weight', 'detailed_description', 'image']

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        if self.instance and self.instance.material:
            self.fields['bin'].initial = self.instance.material.bin

    def save(self, commit=True):
        spec = super().save(commit=False)
        if commit:
            spec.save()
            if spec.material:
                spec.material.bin = self.cleaned_data['bin']
                spec.material.save()
        return spec

class MaterialImageUploadForm(forms.ModelForm):
    class Meta:
        model = MaterialImage
        fields = ['image']
        widgets = {
            'image': forms.ClearableFileInput(),
        }
