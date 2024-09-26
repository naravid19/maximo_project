from django.core.management.base import BaseCommand
from django.db import connection

class Command(BaseCommand):
    help = 'Reset the ID sequence for the Unit table'

    def handle(self, *args, **kwargs):
        with connection.cursor() as cursor:
            # ลบลำดับ ID ในตาราง unit สำหรับ SQLite
            cursor.execute("DELETE FROM sqlite_sequence WHERE name='maximo_app_unit'")
            self.stdout.write(self.style.SUCCESS('ID sequence for Unit table has been reset.'))
