import os
import django

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'material_requisition_system.settings')
django.setup()

from django.urls import reverse
names = [
    'requisitions:export_simple_applicant_requisition_excel',
    'requisitions:export_simple_dispatcher_requisition_excel',
    'requisitions:export_single_requisition_excel'
]

for name in names:
    try:
        if 'category' in name:
            url = reverse(name, kwargs={'category': 'test'})
        elif 'pk' in name or 'single' in name:
            url = reverse(name, kwargs={'pk': 1})
        else:
            url = reverse(name)
        print(f"Name: {name} -> URL: {url}")
    except Exception as e:
        print(f"Name: {name} -> Error: {e}")
