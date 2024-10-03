from django.core.management.base import BaseCommand
from django.db import connection

class Command(BaseCommand):
    help = 'Reset the ID sequence for the Unit table'

    def add_arguments(self, parser):
        # เพิ่ม argument เพื่อรับค่า name
        parser.add_argument(
            '--name',
            type=str,
            default='',
            help='Name of the table to reset the ID sequence'
        )

    def handle(self, *args, **kwargs):
        # ดึงค่าพารามิเตอร์ name จาก kwargs
        table_name = kwargs['name']
        
        with connection.cursor() as cursor:
            # ลบลำดับ ID ในตาราง unit สำหรับ SQLite
            cursor.execute(f"DELETE FROM sqlite_sequence WHERE name='{table_name}'")
            self.stdout.write(self.style.SUCCESS(f'ID sequence for {table_name} table has been reset.'))
