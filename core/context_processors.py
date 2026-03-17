from django.contrib.auth.models import Group

def role_context(request):
    from peer_requests.models import PeerRequest
    is_admin = request.user.is_superuser
    is_authenticated = request.user.is_authenticated
    
    pending_peer_requests_count = 0
    if is_authenticated:
        pending_peer_requests_count = PeerRequest.objects.filter(recipient=request.user, status='pending').count()

    is_applicant = request.user.groups.filter(name='申請人員').exists() if is_authenticated else False
    is_material_handler = request.user.groups.filter(name='撥料人員').exists() if is_authenticated else False
    
    return {
        'is_admin': is_admin,
        'is_applicant': is_applicant,
        'is_material_handler': is_material_handler,
        'pending_peer_requests_count': pending_peer_requests_count,
    }
