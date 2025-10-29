from django import forms
from .models import Material

class MaterialForm(forms.ModelForm):
    class Meta:
        model = Material
        fields = ['location', 'bin', 'material_description', 'system_quantity', 'purchaser']
