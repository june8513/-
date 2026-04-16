import os
import django

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'material_requisition_system.settings')
django.setup()

from requisitions.models import SemiFinishedProcessType

# Define new process points
new_points = [
    {'name': '軟體研發部', 'color': '#EC4899'},
    {'name': '待分配', 'color': '#64748B'},
]

print("正在新增投料點到 SemiFinishedProcessType...")

for pt_data in new_points:
    obj, created = SemiFinishedProcessType.objects.get_or_create(
        name=pt_data['name'],
        defaults={'color': pt_data['color']}
    )
    if created:
        print(f"已建立: {pt_data['name']}")
    else:
        print(f"已存在: {pt_data['name']}")
        # Update color even if exists
        obj.color = pt_data['color']
        obj.save()
        print(f"已更新顏色: {pt_data['name']}")

print("完成。")
