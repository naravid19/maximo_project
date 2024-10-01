import os
import pandas as pd
from django.conf import settings
from django.core.management.base import BaseCommand
from maximo_app.models import Site, PlantType

class Command(BaseCommand):
    help = 'Imports plant_code and site_id from an Excel file and creates Many-to-Many relations'

    def add_arguments(self, parser):
        # ตั้งค่า default file path ไปยังไฟล์ Database.xlsx
        default_file_path = os.path.join(settings.BASE_DIR, 'static', 'excel', 'Database.xlsx')
        
        parser.add_argument(
            'file_path', 
            type=str, 
            nargs='?',  # ทำให้ file_path เป็น optional
            default=default_file_path,  # ใช้ default file path ที่กำหนดไว้
            help='The path to the Excel file'
        )
        parser.add_argument(
            '--sheet_name', 
            type=str, 
            default='Site-PlantType',  # กำหนดค่าเริ่มต้นของ sheet_name
            help='The name of the sheet to read from the Excel file'
        )

    def handle(self, *args, **kwargs):
        file_path = kwargs['file_path']
        sheet_name = kwargs['sheet_name']
        
        # อ่านข้อมูลจากไฟล์ Excel พร้อมกับระบุ sheet_name
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
        except Exception as e:
            self.stdout.write(self.style.ERROR(f"Error reading Excel file: {e}"))
            return
        
        # ลูปผ่านแต่ละแถวและสร้างความสัมพันธ์ Many-to-Many
        for index, row in df.iterrows():
            plant_code = row['plant_code']
            site_id = row['site_id']

            try:
                # ดึง PlantType และ Site จากฐานข้อมูล
                plant_type = PlantType.objects.get(plant_code=plant_code)
                site = Site.objects.get(site_id=site_id)
                
                # เพิ่มความสัมพันธ์ Many-to-Many
                site.plant_types.add(plant_type)

                # บันทึกการเปลี่ยนแปลง
                site.save()
                self.stdout.write(self.style.SUCCESS(f"Successfully added {site_id} to {plant_code}"))
            
            except PlantType.DoesNotExist:
                self.stdout.write(self.style.ERROR(f"PlantType with code {plant_code} does not exist."))
            except Site.DoesNotExist:
                self.stdout.write(self.style.ERROR(f"Site with id {site_id} does not exist."))
