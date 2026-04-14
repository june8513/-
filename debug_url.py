import os
import django

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'material_requisition_system.settings')
django.setup()

from django.urls import reverse
try:
    url = reverse('requisitions:export_single_requisition_excel', kwargs={'pk': 1})
    print(f"URL: {url}")
except Exception as e:
    print(f"Error: {e}")
