from django.contrib.auth.models import Group

def role_context(request):
    from peer_requests.models import PeerRequest
    from requisitions.constants import GROUP_NAMES
    
    is_admin = request.user.is_superuser
    is_authenticated = request.user.is_authenticated
    
    pending_peer_requests_count = 0
    if is_authenticated:
        pending_peer_requests_count = PeerRequest.objects.filter(recipient=request.user, status='pending').count()

    is_applicant = request.user.groups.filter(name='申請人員').exists() if is_authenticated else False
    is_material_handler = request.user.groups.filter(name='撥料人員').exists() if is_authenticated else False
    
    is_applicant_supervisor = request.user.groups.filter(name=GROUP_NAMES['APPLICANT_SUPERVISOR']).exists() if is_authenticated else False
    is_dispatcher_supervisor = request.user.groups.filter(name=GROUP_NAMES['DISPATCHER_SUPERVISOR']).exists() if is_authenticated else False
    
    # 開放主管與管理員存取簡易畫面 (控制導覽列的顯示)
    is_simple_applicant = False
    is_simple_dispatcher = False
    
    if is_authenticated:
        is_simple_applicant = request.user.groups.filter(name=GROUP_NAMES['APPLICANT']).exists() or is_applicant_supervisor or is_admin
        is_simple_dispatcher = request.user.groups.filter(name=GROUP_NAMES['DISPATCHER']).exists() or is_dispatcher_supervisor or is_admin
    
    # 是否為任意主管
    is_supervisor = is_applicant_supervisor or is_dispatcher_supervisor
    
    return {
        'is_admin': is_admin,
        'is_applicant': is_applicant,
        'is_material_handler': is_material_handler,
        'pending_peer_requests_count': pending_peer_requests_count,
        'is_simple_applicant': is_simple_applicant,
        'is_simple_dispatcher': is_simple_dispatcher,
        'is_applicant_supervisor': is_applicant_supervisor,
        'is_dispatcher_supervisor': is_dispatcher_supervisor,
        'is_supervisor': is_supervisor,
    }
