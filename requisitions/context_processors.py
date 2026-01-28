from .constants import GROUP_NAMES


def role_context(request):
    """提供角色相關的上下文變數"""
    is_admin = False
    is_applicant = False
    is_material_handler = False
    is_applicant_supervisor = False
    is_dispatcher_supervisor = False
    is_simple_applicant = False
    is_simple_dispatcher = False

    if request.user.is_authenticated:
        is_admin = request.user.is_superuser
        
        # 檢查主管群組
        is_applicant_supervisor = request.user.groups.filter(
            name=GROUP_NAMES['APPLICANT_SUPERVISOR']
        ).exists()
        is_dispatcher_supervisor = request.user.groups.filter(
            name=GROUP_NAMES['DISPATCHER_SUPERVISOR']
        ).exists()
        
        # 檢查一般人員群組
        is_simple_applicant = request.user.groups.filter(
            name=GROUP_NAMES['APPLICANT']
        ).exists() and not is_applicant_supervisor
        is_simple_dispatcher = request.user.groups.filter(
            name=GROUP_NAMES['DISPATCHER']
        ).exists() and not is_dispatcher_supervisor
        
        # 相容舊版：任何申請/撥料人員（主管或一般）都設為 True
        is_applicant = is_applicant_supervisor or is_simple_applicant or \
                       request.user.groups.filter(name='申請人員').exists()
        is_material_handler = is_dispatcher_supervisor or is_simple_dispatcher or \
                              request.user.groups.filter(name='撥料人員').exists()

    # 是否為任意主管
    is_supervisor = is_applicant_supervisor or is_dispatcher_supervisor

    return {
        'is_admin': is_admin,
        'is_applicant': is_applicant,
        'is_material_handler': is_material_handler,
        'is_applicant_supervisor': is_applicant_supervisor,
        'is_dispatcher_supervisor': is_dispatcher_supervisor,
        'is_simple_applicant': is_simple_applicant,
        'is_simple_dispatcher': is_simple_dispatcher,
        'is_supervisor': is_supervisor,
    }

