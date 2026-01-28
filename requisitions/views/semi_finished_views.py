from django.shortcuts import render, redirect
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.http import JsonResponse
from ..models import SemiFinishedProcessType


def is_supervisor(user):
    """檢查是否為主管"""
    return user.groups.filter(name__in=['申請人員主管', '撥料人員主管', '管理員']).exists() or user.is_superuser


@login_required
def semi_finished_process_type_list(request):
    """半成品投料點管理頁面"""
    if not is_supervisor(request.user):
        messages.error(request, "您沒有權限存取此頁面。")
        return redirect('core:homepage')
    
    process_types = SemiFinishedProcessType.objects.all()
    
    if request.method == 'POST':
        action = request.POST.get('action')
        
        if action == 'create':
            name = request.POST.get('name', '').strip()
            color = request.POST.get('color', '#6366F1')
            if name:
                if SemiFinishedProcessType.objects.filter(name=name).exists():
                    messages.error(request, f"投料點「{name}」已存在。")
                else:
                    SemiFinishedProcessType.objects.create(
                        name=name,
                        color=color,
                        created_by=request.user
                    )
                    messages.success(request, f"已新增投料點「{name}」")
            else:
                messages.error(request, "請輸入投料點名稱。")
        
        elif action == 'update':
            pt_id = request.POST.get('pt_id')
            name = request.POST.get('name', '').strip()
            color = request.POST.get('color', '#6366F1')
            is_active = request.POST.get('is_active') == 'on'
            if pt_id and name:
                SemiFinishedProcessType.objects.filter(pk=pt_id).update(
                    name=name,
                    color=color,
                    is_active=is_active
                )
                messages.success(request, f"已更新投料點「{name}」")
        
        elif action == 'delete':
            pt_id = request.POST.get('pt_id')
            if pt_id:
                pt = SemiFinishedProcessType.objects.filter(pk=pt_id).first()
                if pt:
                    pt.delete()
                    messages.success(request, f"已刪除投料點「{pt.name}」")
        
        return redirect('requisitions:semi_finished_process_type_list')
    
    context = {
        'process_types': process_types,
    }
    return render(request, 'requisitions/semi_finished_process_type_list.html', context)
