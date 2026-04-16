import os
import django

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'material_requisition_system.settings')
django.setup()

from requisitions.models import ProcessType, SemiFinishedProcessType, MachineModel

print('--- ProcessType ---')
for pt in ProcessType.objects.all():
    print(f"{pt.name} (Machine: {pt.machine_model.name})")

print('\n--- SemiFinishedProcessType ---')
for sfpt in SemiFinishedProcessType.objects.all():
    print(sfpt.name)

print('\n--- MachineModels ---')
for mm in MachineModel.objects.all():
    print(mm.name)
