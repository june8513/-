from django.core.management.base import BaseCommand
from django.contrib.auth.models import Group, User
from requisitions.constants import GROUP_NAMES

class Command(BaseCommand):
    help = 'Creates default user groups and assigns superuser to Admin group.'

    def handle(self, *args, **options):
        # Create groups
        groups_to_create = [
            GROUP_NAMES['ADMIN'],
            GROUP_NAMES['APPLICANT_SUPERVISOR'],
            GROUP_NAMES['APPLICANT'],
            GROUP_NAMES['DISPATCHER_SUPERVISOR'],
            GROUP_NAMES['DISPATCHER'],
        ]
        
        created_groups = []
        for group_name in groups_to_create:
            group, created = Group.objects.get_or_create(name=group_name)
            if created:
                self.stdout.write(self.style.SUCCESS(f'Successfully created group: {group_name}'))
                created_groups.append(group_name)
            else:
                self.stdout.write(self.style.WARNING(f'Group already exists: {group_name}'))

        # Assign superuser to Admin group
        try:
            admin_group = Group.objects.get(name=GROUP_NAMES['ADMIN'])
            superuser = User.objects.filter(is_superuser=True).first()
            if superuser:
                superuser.groups.add(admin_group)
                self.stdout.write(self.style.SUCCESS(f'Successfully assigned superuser "{superuser.username}" to group "{GROUP_NAMES["ADMIN"]}"'))
            else:
                self.stdout.write(self.style.WARNING('No superuser found. Please create one using "python manage.py createsuperuser"'))
        except Exception as e:
            self.stdout.write(self.style.ERROR(f'Error assigning superuser to group: {e}'))
        
        self.stdout.write(self.style.SUCCESS(f'\nTotal groups created: {len(created_groups)}'))
        self.stdout.write(self.style.SUCCESS('Setup roles completed!'))
