from django import forms
from .models import MaterialSpecification

class MaterialSpecificationForm(forms.ModelForm):
    class Meta:
        model = MaterialSpecification
        fields = ['size', 'weight', 'detailed_description', 'image']
