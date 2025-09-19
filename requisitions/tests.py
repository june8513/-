from django.test import TestCase, Client
from django.contrib.auth.models import User, Group
from django.urls import reverse
from django.utils import timezone
from datetime import date
from decimal import Decimal
from django.db.utils import IntegrityError

from .models import Requisition, MaterialListVersion, RequisitionItem

class RequisitionModelTest(TestCase):
    def setUp(self):
        self.user_applicant = User.objects.create_user(username='applicant1', password='password123')
        self.user_material_handler = User.objects.create_user(username='materialhandler1', password='password123')
        self.user_admin = User.objects.create_user(username='admin1', password='password123', is_superuser=True)

        # Create groups if they don't exist
        Group.objects.get_or_create(name='申請人員')
        Group.objects.get_or_create(name='撥料人員')

        self.applicant_group = Group.objects.get(name='申請人員')
        self.material_handler_group = Group.objects.get(name='撥料人員')

        self.user_applicant.groups.add(self.applicant_group)
        self.user_material_handler.groups.add(self.material_handler_group)

    def test_requisition_creation(self):
        """Test Requisition model creation and __str__ method."""
        requisition = Requisition.objects.create(
            work_order_number='WO12345',
            applicant=self.user_applicant,
            request_date=date.today(),
            process_type='machine_head',
            status='pending'
        )
        self.assertEqual(requisition.work_order_number, 'WO12345')
        self.assertEqual(requisition.applicant, self.user_applicant)
        self.assertEqual(requisition.request_date, date.today())
        self.assertEqual(requisition.process_type, 'machine_head')
        self.assertEqual(requisition.status, 'pending')
        self.assertIsNotNone(requisition.created_at)
        self.assertIsNotNone(requisition.updated_at)
        self.assertEqual(str(requisition), f"撥料申請單: WO12345 (機頭) - {self.user_applicant.username}")

    def test_requisition_unique_constraint_same_work_order_and_process_type(self):
        """Test that work_order_number and process_type are unique together."""
        Requisition.objects.create(
            work_order_number='WO12345',
            applicant=self.user_applicant,
            request_date=date.today(),
            process_type='machine_head',
            status='pending'
        )
        with self.assertRaises(IntegrityError):
            Requisition.objects.create(
                work_order_number='WO12345',
                applicant=self.user_applicant,
                request_date=date.today(),
                process_type='machine_head', # Same work_order_number and process_type
                status='pending'
            )

    def test_requisition_unique_constraint_different_process_type(self):
        """Test that different process_type with same work_order_number is allowed."""
        Requisition.objects.create(
            work_order_number='WO12345',
            applicant=self.user_applicant,
            request_date=date.today(),
            process_type='machine_head',
            status='pending'
        )
        requisition2 = Requisition.objects.create(
            work_order_number='WO12345',
            applicant=self.user_applicant,
            request_date=date.today(),
            process_type='spindle', # Different process_type
            status='pending'
        )
        self.assertIsNotNone(requisition2.pk)

    def test_material_list_version_creation(self):
        """Test MaterialListVersion model creation and __str__ method."""
        requisition = Requisition.objects.create(
            work_order_number='WO12346',
            applicant=self.user_applicant,
            request_date=date.today(),
            process_type='electrical',
            status='pending'
        )
        material_version = MaterialListVersion.objects.create(
            requisition=requisition,
            uploaded_by=self.user_material_handler
        )
        self.assertEqual(material_version.requisition, requisition)
        self.assertEqual(material_version.uploaded_by, self.user_material_handler)
        self.assertIsNotNone(material_version.uploaded_at)
        self.assertIn(requisition.work_order_number, str(material_version))

    def test_requisition_item_creation(self):
        """Test RequisitionItem model creation and __str__ method."""
        requisition = Requisition.objects.create(
            work_order_number='WO12347',
            applicant=self.user_applicant,
            request_date=date.today(),
            process_type='system',
            status='pending'
        )
        material_version = MaterialListVersion.objects.create(
            requisition=requisition,
            uploaded_by=self.user_material_handler
        )
        item = RequisitionItem.objects.create(
            material_list_version=material_version,
            work_order_number='WO12347',
            item_number='ITEM001',
            item_name='Test Item',
            required_quantity=Decimal('10.50'),
            stock_quantity=Decimal('5.00'),
            confirmed_quantity=Decimal('5.00'),
            is_signed_off=False
        )
        self.assertEqual(item.material_list_version, material_version)
        self.assertEqual(item.work_order_number, 'WO12347')
        self.assertEqual(item.item_number, 'ITEM001')
        self.assertEqual(item.item_name, 'Test Item')
        self.assertEqual(item.required_quantity, Decimal('10.50'))
        self.assertEqual(item.stock_quantity, Decimal('5.00'))
        self.assertEqual(item.confirmed_quantity, Decimal('5.00'))
        self.assertFalse(item.is_signed_off)
        self.assertEqual(str(item), "Test Item (10.50)")


class UserAuthenticationTest(TestCase):
    def setUp(self):
        self.client = Client()
        self.user = User.objects.create_user(username='testuser', password='testpassword')

    def test_login_success(self):
        """Test successful user login."""
        response = self.client.post(reverse('login'), {'username': 'testuser', 'password': 'testpassword'})
        self.assertRedirects(response, reverse('requisition_list'))
        self.assertTrue(self.client.session.get('_auth_user_id'))

    def test_login_invalid_credentials(self):
        """Test user login with invalid credentials."""
        response = self.client.post(reverse('login'), {'username': 'testuser', 'password': 'wrongpassword'})
        self.assertContains(response, "無效的使用者名稱或密碼。")
        self.assertFalse(self.client.session.get('_auth_user_id'))

    def test_logout(self):
        """Test user logout."""
        self.client.login(username='testuser', password='testpassword')
        response = self.client.get(reverse('logout'))
        self.assertRedirects(response, reverse('login'))
        self.assertFalse(self.client.session.get('_auth_user_id'))


class RequisitionCreateViewTest(TestCase):
    def setUp(self):
        self.client = Client()
        self.applicant_user = User.objects.create_user(username='applicant', password='password')
        self.admin_user = User.objects.create_user(username='admin', password='password', is_superuser=True)
        self.material_handler_user = User.objects.create_user(username='material_handler', password='password')

        Group.objects.get_or_create(name='申請人員')
        Group.objects.get_or_create(name='撥料人員')

        self.applicant_group = Group.objects.get(name='申請人員')
        self.material_handler_group = Group.objects.get(name='撥料人員')

        self.applicant_user.groups.add(self.applicant_group)
        self.material_handler_user.groups.add(self.material_handler_group)

    def test_requisition_create_unauthorized(self):
        """Test that unauthorized users cannot access requisition_create view."""
        response = self.client.get(reverse('requisition_create'))
        self.assertRedirects(response, '/accounts/login/?next=/requisitions/create/') # Redirects to login

        self.client.login(username='material_handler', password='password')
        response = self.client.get(reverse('requisition_create'))
        self.assertRedirects(response, reverse('requisition_list')) # Redirects to list with error message
        self.assertContains(response, "您沒有權限建立撥料申請單。", html=True)

    def test_requisition_create_by_applicant(self):
        """Test that an applicant can create a requisition."""
        self.client.login(username='applicant', password='password')
        response = self.client.post(reverse('requisition_create'), {
            'work_order_number': 'WO_APP_001',
            'request_date': date.today(),
            'process_type': 'machine_head',
        })
        self.assertRedirects(response, reverse('requisition_list'))
        self.assertTrue(Requisition.objects.filter(work_order_number='WO_APP_001', applicant=self.applicant_user).exists())
        self.assertContains(response, "撥料申請單建立成功！", html=True)

    def test_requisition_create_by_admin(self):
        """Test that an admin can create a requisition."""
        self.client.login(username='admin', password='password')
        response = self.client.post(reverse('requisition_create'), {
            'work_order_number': 'WO_ADMIN_001',
            'request_date': date.today(),
            'process_type': 'spindle',
        })
        self.assertRedirects(response, reverse('requisition_list'))
        self.assertTrue(Requisition.objects.filter(work_order_number='WO_ADMIN_001', applicant=self.admin_user).exists())
        self.assertContains(response, "撥料申請單建立成功！", html=True)

    def test_requisition_create_duplicate_work_order_process_type(self):
        """Test that creating a duplicate work order number for the same process type fails."""
        self.client.login(username='applicant', password='password')
        Requisition.objects.create(
            work_order_number='DUPLICATE_WO',
            applicant=self.applicant_user,
            request_date=date.today(),
            process_type='machine_head',
            status='pending'
        )
        response = self.client.post(reverse('requisition_create'), {
            'work_order_number': 'DUPLICATE_WO',
            'request_date': date.today(),
            'process_type': 'machine_head',
        })
        self.assertContains(response, "此工單單號在該需求流程中已存在，請使用不同的工單單號或需求流程。", html=True)
        self.assertEqual(Requisition.objects.filter(work_order_number='DUPLICATE_WO', process_type='machine_head').count(), 1)

    def test_requisition_create_different_process_type_same_work_order(self):
        """Test that creating same work order number with different process type succeeds."""
        self.client.login(username='applicant', password='password')
        Requisition.objects.create(
            work_order_number='SAME_WO',
            applicant=self.applicant_user,
            request_date=date.today(),
            process_type='machine_head',
            status='pending'
        )
        response = self.client.post(reverse('requisition_create'), {
            'work_order_number': 'SAME_WO',
            'request_date': date.today(),
            'process_type': 'spindle',
        })
        self.assertRedirects(response, reverse('requisition_list'))
        self.assertTrue(Requisition.objects.filter(work_order_number='SAME_WO', process_type='spindle').exists())
        self.assertEqual(Requisition.objects.filter(work_order_number='SAME_WO').count(), 2)