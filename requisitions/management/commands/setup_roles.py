from django.core.management.base import BaseCommand
from django.contrib.auth.models import Group, User

class Command(BaseCommand):
    help = 'Creates default user groups and assigns superuser to Admin group.'

    def handle(self, *args, **options):
        # Create groups
        admin_group, created = Group.objects.get_or_create(name='管理員')
        if created:
            self.stdout.write(self.style.SUCCESS('Successfully created group: 管理員'))
        
        applicant_group, created = Group.objects.get_or_create(name='申請人員')
        if created:
            self.stdout.write(self.style.SUCCESS('Successfully created group: 申請人員'))

        material_handler_group, created = Group.objects.get_or_create(name='撥料人員')
        if created:
            self.stdout.write(self.style.SUCCESS('Successfully created group: 撥料人員'))

        # Assign superuser to Admin group
        try:
            superuser = User.objects.filter(is_superuser=True).first()
            if superuser:
                superuser.groups.add(admin_group)
                self.stdout.write(self.style.SUCCESS(f'Successfully assigned superuser "{superuser.username}" to group "管理員"'))
            else:
                self.stdout.write(self.style.WARNING('No superuser found. Please create one using "python manage.py createsuperuser"'))
        except Exception as e:
            self.stdout.write(self.style.ERROR(f'Error assigning superuser to group: {e}'))
