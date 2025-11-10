from django.contrib.auth.models import Group

def role_context(request):
    is_admin = request.user.is_superuser
    is_applicant = request.user.groups.filter(name='申請人員').exists()
    is_material_handler = request.user.groups.filter(name='撥料人員').exists()
    return {
        'is_admin': is_admin,
        'is_applicant': is_applicant,
        'is_material_handler': is_material_handler,
    }
