from django import forms
from .models import MaterialSpecification
from inventory.models import MaterialImage

class MaterialSpecificationForm(forms.ModelForm):
    class Meta:
        model = MaterialSpecification
        fields = ['size', 'weight', 'detailed_description', 'image']

class MaterialImageUploadForm(forms.ModelForm):
    class Meta:
        model = MaterialImage
        fields = ['image']
        widgets = {
            'image': forms.ClearableFileInput(),
        }
