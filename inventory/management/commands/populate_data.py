from django.core.management.base import BaseCommand
from django.contrib.auth.models import User
from inventory.models import Material, MaterialTransaction, StorageLocation
from requisitions.models import WorkOrderMaterial, Requisition, ProcessType, MachineModel # Import necessary models
from django.utils import timezone
from datetime import timedelta
import random

class Command(BaseCommand):
    help = 'Populates the database with sample data for testing.'

    def handle(self, *args, **options):
        self.stdout.write(self.style.SUCCESS('Starting data population...'))

        # --- Get or Create User ---
        user, _ = User.objects.get_or_create(username='testuser', defaults={'email': 'test@example.com', 'password': 'password'})

        # --- Get or Create Prerequisite objects for Requisition and WorkOrderMaterial ---
        machine_model, _ = MachineModel.objects.get_or_create(name='Default Model')
        process_type, _ = ProcessType.objects.get_or_create(name='Default Process', machine_model=machine_model)
        requisition, _ = Requisition.objects.get_or_create(
            order_number='TEST-ORDER-001', 
            process_type=process_type.name, 
            defaults={'applicant': user, 'request_date': timezone.now().date()}
        )

        # --- Get or Create Materials ---
        location, _ = StorageLocation.objects.get_or_create(name='主倉庫')
        material1, _ = Material.objects.get_or_create(material_code='MAT001', defaults={'material_description':'螺絲 M5', 'system_quantity':1000, 'location':location, 'bin':'A01'})
        material2, _ = Material.objects.get_or_create(material_code='MAT002', defaults={'material_description':'電線 1mm', 'system_quantity':500, 'location':location, 'bin':'A02'})
        material3, _ = Material.objects.get_or_create(material_code='MAT003', defaults={'material_description':'螺帽 M5', 'system_quantity':800, 'location':location, 'bin':'A03'})

        # --- Clear and Create MaterialTransaction data ---
        MaterialTransaction.objects.all().delete()
        self.stdout.write(self.style.SUCCESS('Cleared existing MaterialTransaction data.'))
        transactions_to_create = [
            MaterialTransaction(material=material1, user=user, transaction_type='ALLOCATION', quantity_change=50, new_system_quantity=950, timestamp=timezone.now()),
            MaterialTransaction(material=material2, user=user, transaction_type='ALLOCATION', quantity_change=100, new_system_quantity=400, timestamp=timezone.now()),
        ]
        MaterialTransaction.objects.bulk_create(transactions_to_create)
        self.stdout.write(self.style.SUCCESS(f'Successfully created {len(transactions_to_create)} sample MaterialTransaction records.'))

        # --- Clear and Create WorkOrderMaterial data ---
        WorkOrderMaterial.objects.all().delete()
        self.stdout.write(self.style.SUCCESS('Cleared existing WorkOrderMaterial data.'))
        work_orders_to_create = [
            # One for today
            WorkOrderMaterial(machine_model=machine_model, order_number='TEST-ORDER-001', material_number=material1.material_code, item_name='螺絲 M5', required_quantity=200, process_type=process_type, estimated_arrival_date=timezone.now().date()),
            # One for yesterday
            WorkOrderMaterial(machine_model=machine_model, order_number='TEST-ORDER-002', material_number=material2.material_code, item_name='電線 1mm', required_quantity=300, process_type=process_type, estimated_arrival_date=timezone.now().date() - timedelta(days=1)),
            # Another one for today
            WorkOrderMaterial(machine_model=machine_model, order_number='TEST-ORDER-003', material_number=material3.material_code, item_name='螺帽 M5', required_quantity=400, process_type=process_type, estimated_arrival_date=timezone.now().date()),
        ]
        WorkOrderMaterial.objects.bulk_create(work_orders_to_create)
        self.stdout.write(self.style.SUCCESS(f'Successfully created {len(work_orders_to_create)} sample WorkOrderMaterial records.'))

        self.stdout.write(self.style.SUCCESS('Data population complete.'))
