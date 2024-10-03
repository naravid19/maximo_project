from django.core.management.base import BaseCommand
from django.db import connection

class Command(BaseCommand):
    help = (
        'Reset the ID sequence for the specified table (SQLite database only).\n\n'
        'Usage:\n'
        '  python manage.py reset_id_sequence --name=<table_name>\n\n'
        'Example:\n'
        '  python manage.py reset_id_sequence --name=maximo_app_site\n'
    )

    def add_arguments(self, parser):
        # เพิ่ม argument เพื่อรับค่า name
        parser.add_argument(
            '--name',
            type=str,
            required=True,
            help='Name of the table to reset the ID sequence (e.g., unit)'
        )

    def handle(self, *args, **kwargs):
        # ดึงค่าพารามิเตอร์ name จาก kwargs
        table_name = kwargs['name']
        
        if not table_name:
            self.stdout.write(self.style.ERROR('Please specify the table name using --name=<table_name>.'))
            return
        
        with connection.cursor() as cursor:
            # ลบลำดับ ID ในตาราง unit สำหรับ SQLite
            try:
                cursor.execute(f"DELETE FROM sqlite_sequence WHERE name='{table_name}'")
                self.stdout.write(self.style.SUCCESS(f'The ID sequence for the {table_name} table has been reset.'))
            except Exception as e:
                self.stdout.write(self.style.ERROR(f'Error: {e}'))
