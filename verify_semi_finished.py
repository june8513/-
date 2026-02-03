import os
import django

# Setup Django first - strictly before any other django imports
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'material_requisition_system.settings')
django.setup()

# Now import Django components
from django.test import RequestFactory
from django.contrib.auth.models import User, Group
from requisitions.views.simple_views import simple_dispatcher_home, simple_applicant_detail
from requisitions.models import Requisition
from requisitions.constants import GROUP_NAMES

def setup_test_data():
    # Ensure simplified groups exist
    dispatcher_group, _ = Group.objects.get_or_create(name=GROUP_NAMES['DISPATCHER'])
    applicant_group, _ = Group.objects.get_or_create(name=GROUP_NAMES['APPLICANT'])

    # Create or get a test dispatcher
    dispatcher, _ = User.objects.get_or_create(username='test_dispatcher')
    dispatcher.set_password('password')
    dispatcher.groups.add(dispatcher_group)
    dispatcher.save()

    # Create or get a test applicant
    applicant, _ = User.objects.get_or_create(username='test_applicant')
    applicant.set_password('password')
    applicant.groups.add(applicant_group)
    applicant.save()
    
    return dispatcher, applicant

def verify_semi_finished_logic():
    print("Starting verification...")
    dispatcher, applicant = setup_test_data()
    factory = RequestFactory()

    # Check for existing semi-finished requisition
    req = Requisition.objects.filter(requisition_type='semi_finished', applicant=applicant).first()
    
    if not req:
        print("No semi-finished requisition found for test_applicant. Creating one...")
        req = Requisition.objects.create(
            applicant=applicant,
            order_number='TEST-SEMI-001',
            requisition_type='semi_finished',
            status='demand_submitted',
            process_type='SEMI',
            request_date=django.utils.timezone.now().date()
        )
    else:
        req.status = 'demand_submitted'
        req.save()
        print(f"Using existing requisition: {req.pk}")

    # --- Test 1: Simple Dispatcher Home (Semi-Finished) ---
    print("\n[Test 1] Testing simple_dispatcher_home (type=semi_finished)...")
    request = factory.get('/requisitions/simple/dispatcher/', {'type': 'semi_finished'})
    request.user = dispatcher
    
    response = simple_dispatcher_home(request)
    content = response.content.decode('utf-8')
    
    # Needs to find the applicant's username in the response (because we group by applicant)
    if applicant.username in content:
        print("✅ SUCCESS: Found applicant username in dispatcher home.")
    else:
        print(f"❌ FAILURE: Did not find '{applicant.username}' in dispatcher home.")

    # --- Test 2: Simple Applicant Detail ---
    print("\n[Test 2] Testing simple_applicant_detail...")
    request = factory.get(f'/requisitions/simple/applicant/{req.pk}/')
    request.user = applicant
    
    response = simple_applicant_detail(request, req.pk)
    content = response.content.decode('utf-8')
    
    # Should display "領料人"
    if "領料人" in content:
        print("✅ SUCCESS: Found '領料人' label in detail view.")
    else:
        print("❌ FAILURE: Detail view missing '領料人' label.")
        
    if applicant.username in content:
        print("✅ SUCCESS: Found applicant name in detail view.")
    else:
        print("❌ FAILURE: Detail view missing applicant name.")

    print("\nVerification complete.")

if __name__ == "__main__":
    verify_semi_finished_logic()
