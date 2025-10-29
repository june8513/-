from django.test import TestCase, Client
from django.contrib.auth.models import User, Group
from django.urls import reverse
from django.utils import timezone
from datetime import date
from decimal import Decimal
from django.db.utils import IntegrityError

from .models import Requisition, MaterialListVersion, RequisitionItem, MachineModel, ProcessType, WorkOrderMaterial

class RequisitionModelTest(TestCase):
    @classmethod
    def setUpTestData(cls):
        # Create groups if they don't exist
        applicant_group, _ = Group.objects.get_or_create(name='申請人員')
        material_handler_group, _ = Group.objects.get_or_create(name='撥料人員')

        cls.user_applicant = User.objects.create_user(username='applicant1', password='password123')
        cls.user_material_handler = User.objects.create_user(username='materialhandler1', password='password123')
        cls.user_admin = User.objects.create_user(username='admin1', password='password123', is_superuser=True)

        cls.user_applicant.groups.add(applicant_group)
        cls.user_material_handler.groups.add(material_handler_group)

        # Create MachineModel and ProcessType instances for tests
        cls.machine_model_head, _ = MachineModel.objects.get_or_create(name='機頭')
        cls.machine_model_spindle, _ = MachineModel.objects.get_or_create(name='主軸')
        cls.machine_model_electrical, _ = MachineModel.objects.get_or_create(name='電裝')
        cls.machine_model_system, _ = MachineModel.objects.get_or_create(name='機械')

        cls.process_type_head, _ = ProcessType.objects.get_or_create(name='機頭', machine_model=cls.machine_model_head)
        cls.process_type_spindle, _ = ProcessType.objects.get_or_create(name='主軸', machine_model=cls.machine_model_spindle)
        cls.process_type_electrical, _ = ProcessType.objects.get_or_create(name='電裝', machine_model=cls.machine_model_electrical)
        cls.process_type_system, _ = ProcessType.objects.get_or_create(name='機械', machine_model=cls.machine_model_system)

    def setUp(self):
        pass # No need for setUp in this class anymore, as data is created in setUpTestData

    def test_requisition_creation(self):
        """Test Requisition model creation and __str__ method."""
        requisition = Requisition.objects.create(
            order_number='WO12345',
            applicant=self.user_applicant,
            request_date=date.today(),
            process_type='機頭',
            status='pending'
        )
        self.assertEqual(requisition.order_number, 'WO12345')
        self.assertEqual(requisition.applicant, self.user_applicant)
        self.assertEqual(requisition.request_date, date.today())
        self.assertEqual(requisition.process_type, '機頭')
        self.assertEqual(requisition.status, 'pending')
        self.assertIsNotNone(requisition.created_at)
        self.assertIsNotNone(requisition.updated_at)
        self.assertEqual(str(requisition), f"撥料申請單: WO12345 (機頭) - {self.user_applicant.username}")

    def test_requisition_unique_constraint_same_work_order_and_process_type(self):
        """Test that work_order_number and process_type are unique together."""
        Requisition.objects.create(
            order_number='WO12345',
            applicant=self.user_applicant,
            request_date=date.today(),
            process_type='機頭',
            status='pending'
        )
        with self.assertRaises(IntegrityError):
            Requisition.objects.create(
                order_number='WO12345',
                applicant=self.user_applicant,
                request_date=date.today(),
                process_type='機頭', # Same work_order_number and process_type
                status='pending'
            )

    def test_requisition_unique_constraint_different_process_type(self):
        """Test that different process_type with same work_order_number is allowed."""
        Requisition.objects.create(
            order_number='WO12345',
            applicant=self.user_applicant,
            request_date=date.today(),
            process_type='機頭',
            status='pending'
        )
        requisition2 = Requisition.objects.create(
            order_number='WO12345',
            applicant=self.user_applicant,
            request_date=date.today(),
            process_type='主軸', # Different process_type
            status='pending'
        )
        self.assertIsNotNone(requisition2.pk)

    def test_material_list_version_creation(self):
        """Test MaterialListVersion model creation and __str__ method."""
        requisition = Requisition.objects.create(
            order_number='WO12346',
            applicant=self.user_applicant,
            request_date=date.today(),
            process_type='電裝',
            status='pending'
        )
        material_version = MaterialListVersion.objects.create(
            requisition=requisition,
            uploaded_by=self.user_material_handler
        )
        self.assertEqual(material_version.requisition, requisition)
        self.assertEqual(material_version.uploaded_by, self.user_material_handler)
        self.assertIsNotNone(material_version.uploaded_at)
        self.assertIn(requisition.order_number, str(material_version))

    def test_requisition_item_creation(self):
        """Test RequisitionItem model creation and __str__ method."""
        requisition = Requisition.objects.create(
            order_number='WO12347',
            applicant=self.user_applicant,
            request_date=date.today(),
            process_type='機械',
            status='pending'
        )
        material_version = MaterialListVersion.objects.create(
            requisition=requisition,
            uploaded_by=self.user_material_handler
        )
        item = RequisitionItem.objects.create(
            material_list_version=material_version,
            order_number='WO12347',
            material_number='ITEM001',
            item_name='Test Item',
            required_quantity=Decimal('10.50'),
            stock_quantity=Decimal('5.00'),
            confirmed_quantity=Decimal('5.00'),
            is_signed_off=False
        )
        self.assertEqual(item.material_list_version, material_version)
        self.assertEqual(item.order_number, 'WO12347')
        self.assertEqual(item.material_number, 'ITEM001')
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
        self.assertRedirects(response, reverse('homepage'))
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
    @classmethod
    def setUpTestData(cls):
        cls.applicant_user = User.objects.create_user(username='applicant', password='password')
        cls.admin_user = User.objects.create_user(username='admin', password='password', is_superuser=True)
        cls.material_handler_user = User.objects.create_user(username='material_handler', password='password')

        applicant_group, _ = Group.objects.get_or_create(name='申請人員')
        material_handler_group, _ = Group.objects.get_or_create(name='撥料人員')

        cls.applicant_user.groups.add(applicant_group)
        cls.material_handler_user.groups.add(material_handler_group)

        # Create MachineModel and ProcessType instances for tests
        cls.machine_model_head, _ = MachineModel.objects.get_or_create(name='機頭')
        cls.machine_model_spindle, _ = MachineModel.objects.get_or_create(name='主軸')
        cls.machine_model_electrical, _ = MachineModel.objects.get_or_create(name='電裝')
        cls.machine_model_system, _ = MachineModel.objects.get_or_create(name='機械')

        cls.process_type_head, _ = ProcessType.objects.get_or_create(name='機頭', machine_model=cls.machine_model_head)
        cls.process_type_spindle, _ = ProcessType.objects.get_or_create(name='主軸', machine_model=cls.machine_model_spindle)
        cls.process_type_electrical, _ = ProcessType.objects.get_or_create(name='電裝', machine_model=cls.machine_model_electrical)
        cls.process_type_system, _ = ProcessType.objects.get_or_create(name='機械', machine_model=cls.machine_model_system)

        # Create dummy WorkOrderMaterial objects to make process_type choices available
        WorkOrderMaterial.objects.create(
            order_number='WO_APP_001',
            material_number='MAT001',
            item_name='Test Material 1',
            required_quantity=10,
            process_type=cls.process_type_head,
            machine_model=cls.machine_model_head
        )
        WorkOrderMaterial.objects.create(
            order_number='WO_ADMIN_001',
            material_number='MAT002',
            item_name='Test Material 2',
            required_quantity=20,
            process_type=cls.process_type_spindle,
            machine_model=cls.machine_model_spindle
        )
        WorkOrderMaterial.objects.create(
            order_number='SAME_WO_1',
            material_number='MAT003',
            item_name='Test Material 3',
            required_quantity=30,
            process_type=cls.process_type_head,
            machine_model=cls.machine_model_head
        )
        WorkOrderMaterial.objects.create(
            order_number='SAME_WO_2',
            material_number='MAT004',
            item_name='Test Material 4',
            required_quantity=40,
            process_type=cls.process_type_spindle,
            machine_model=cls.machine_model_spindle
        )
        WorkOrderMaterial.objects.create(
            order_number='DUPLICATE_WO_TEST',
            material_number='MAT005',
            item_name='Test Material 5',
            required_quantity=50,
            process_type=cls.process_type_head,
            machine_model=cls.machine_model_head
        )

    def setUp(self):
        self.client = Client()

    def test_requisition_create_unauthorized(self):
        """Test that unauthorized users cannot access requisition_create view."""
        response = self.client.get(reverse('requisition_create'))
        self.assertRedirects(response, '/requisitions/login/?next=/requisitions/create/') # Redirects to login

        self.client.login(username=self.material_handler_user.username, password='password')
        response = self.client.get(reverse('requisition_create'))
        self.assertRedirects(response, reverse('requisition_list')) # Redirects to list with error message

    def test_requisition_create_by_applicant(self):
        """Test that an applicant can create a requisition."""
        self.client.login(username=self.applicant_user.username, password='password')
        response = self.client.post(reverse('requisition_create'), {
            'order_number': 'WO_APP_001',
            'request_date': date.today(),
            'process_type': self.process_type_head.id,
        })
        self.assertRedirects(response, reverse('requisition_list'))
        self.assertTrue(Requisition.objects.filter(order_number='WO_APP_001', applicant=self.applicant_user).exists())
        self.assertContains(response, "撥料申請單建立成功！", html=True)

    def test_requisition_create_by_admin(self):
        """Test that an admin can create a requisition."""
        self.client.login(username=self.admin_user.username, password='password')
        response = self.client.post(reverse('requisition_create'), {
            'order_number': 'WO_ADMIN_001',
            'request_date': date.today(),
            'process_type': self.process_type_spindle.id,
        })
        self.assertRedirects(response, reverse('requisition_list'))
        self.assertTrue(Requisition.objects.filter(order_number='WO_ADMIN_001', applicant=self.admin_user).exists())
        self.assertContains(response, "撥料申請單建立成功！", html=True)

    def test_requisition_create_duplicate_work_order_process_type(self):
        """Test that creating a duplicate work order number for the same process type fails."""
        self.client.login(username=self.applicant_user.username, password='password')
        Requisition.objects.create(
            order_number='DUPLICATE_WO_TEST',
            applicant=self.applicant_user,
            request_date=date.today(),
            process_type='機頭',
            status='pending'
        )
        response = self.client.post(reverse('requisition_create'), {
            'order_number': 'DUPLICATE_WO_TEST',
            'request_date': date.today(),
            'process_type': self.process_type_head.id,
        })
        self.assertContains(response, "此訂單單號在該需求流程中已存在，請選擇不同的訂單單號或需求流程，或修改現有申請單。", html=True)
        self.assertEqual(Requisition.objects.filter(order_number='DUPLICATE_WO_TEST', process_type='機頭').count(), 1)

    def test_requisition_create_different_process_type_same_work_order(self):
        """Test that creating same work order number with different process type succeeds."""
        self.client.login(username=self.applicant_user.username, password='password')
        Requisition.objects.create(
            order_number='SAME_WO_1',
            applicant=self.applicant_user,
            request_date=date.today(),
            process_type='機頭',
            status='pending'
        )
        response = self.client.post(reverse('requisition_create'), {
            'order_number': 'SAME_WO_1',
            'request_date': date.today(),
            'process_type': self.process_type_spindle.id,
        })
        self.assertRedirects(response, reverse('requisition_list'))
        self.assertTrue(Requisition.objects.filter(order_number='SAME_WO_1', process_type='主軸').exists())
        self.assertEqual(Requisition.objects.filter(order_number='SAME_WO_1').count(), 2)
        self.assertContains(response, "撥料申請單建立成功！", html=True)