from django.shortcuts import render, redirect
from requisitions.constants import GROUP_NAMES


def homepage(request):
    """
    首頁視圖 - 根據使用者角色導向不同介面
    """
    if request.user.is_authenticated:
        is_admin = request.user.is_superuser
        
        # 檢查是否為簡易人員（非主管）
        is_simple_applicant = request.user.groups.filter(
            name=GROUP_NAMES['APPLICANT']
        ).exists() and not request.user.groups.filter(
            name=GROUP_NAMES['APPLICANT_SUPERVISOR']
        ).exists() and not is_admin
        
        is_simple_dispatcher = request.user.groups.filter(
            name=GROUP_NAMES['DISPATCHER']
        ).exists() and not request.user.groups.filter(
            name=GROUP_NAMES['DISPATCHER_SUPERVISOR']
        ).exists() and not is_admin
        
        # 簡易人員自動導向簡易介面
        if is_simple_applicant:
            return redirect('requisitions:simple_applicant_home')
        if is_simple_dispatcher:
            return redirect('requisitions:simple_dispatcher_home')
        
        # 主管和管理員使用完整介面
        is_applicant_supervisor = request.user.groups.filter(
            name=GROUP_NAMES['APPLICANT_SUPERVISOR']
        ).exists()
        is_dispatcher_supervisor = request.user.groups.filter(
            name=GROUP_NAMES['DISPATCHER_SUPERVISOR']
        ).exists()
        
        # 相容舊版群組名稱
        is_applicant = is_applicant_supervisor or request.user.groups.filter(name='申請人員').exists()
        is_material_handler = is_dispatcher_supervisor or request.user.groups.filter(name='撥料人員').exists()
        
        context = {
            'is_admin': is_admin,
            'is_applicant': is_applicant,
            'is_material_handler': is_material_handler,
            'is_applicant_supervisor': is_applicant_supervisor,
            'is_dispatcher_supervisor': is_dispatcher_supervisor,
        }
        return render(request, 'core/homepage.html', context)
    else:
        return render(request, 'core/landing.html')