from django.apps import AppConfig


class RequisitionsConfig(AppConfig):
    default_auto_field = 'django.db.models.BigAutoField'
    name = 'requisitions'

    def ready(self):
        print("DEBUG: RequisitionsConfig.ready() is being executed. Importing signals...") # DEBUG
        import requisitions.signals  # noqa
        print("DEBUG: Requisitions signals imported.") # DEBUG
