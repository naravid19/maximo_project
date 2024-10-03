import os
import pandas as pd
import logging
from django.conf import settings
from django.core.management.base import BaseCommand
from maximo_app.models import Site, PlantType
from django.db import transaction

# ตั้งค่า log สำหรับบันทึกข้อความการทำงาน
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class Command(BaseCommand):
    help = 'Imports PLANT_CODE and SITE_ID from an Excel file and creates Many-to-Many relations'

    def handle(self, *args, **kwargs):
        # กำหนดเส้นทางไปยังไฟล์ Database.xlsx และชื่อ sheet ที่ต้องการอ่าน
        file_path = os.path.join(settings.BASE_DIR, 'static', 'excel', 'Database.xlsx')
        sheet_name = 'plant_site'

        # ตรวจสอบว่าไฟล์ Excel มีอยู่หรือไม่
        if not os.path.exists(file_path):
            logger.error(f"File {file_path} not found")
            self.stdout.write(self.style.ERROR(f"File {file_path} not found"))
            return
        
        try:
            # อ่านข้อมูลจาก Excel
            df = pd.read_excel(file_path, sheet_name=sheet_name)
        except FileNotFoundError:
            logger.error("Error: The specified file was not found.")
            self.stdout.write(self.style.ERROR("Error: The specified file was not found."))
            return
        except pd.errors.EmptyDataError:
            logger.error("Error: The Excel file is empty.")
            self.stdout.write(self.style.ERROR("Error: The Excel file is empty."))
            return
        except Exception as e:
            logger.error(f"Error reading the Excel file: {str(e)}")
            self.stdout.write(self.style.ERROR(f"Error reading the Excel file: {str(e)}"))
            return

        # ตรวจสอบว่าไฟล์ Excel มีข้อมูลหรือไม่
        if df.empty:
            logger.error("The Excel file is empty. No data to import.")
            self.stdout.write(self.style.ERROR("The Excel file is empty. No data to import."))
            return
        
        df.columns = df.columns.str.upper()
        
        # ตรวจสอบว่าคอลัมน์ที่จำเป็นมีอยู่ใน DataFrame หรือไม่
        required_columns = ['PLANT_CODE', 'SITE_ID']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            logger.error(f"Missing required columns in the Excel file: {', '.join(missing_columns)}")
            self.stdout.write(self.style.ERROR(f"Missing required columns in the Excel file: {', '.join(missing_columns)}"))
            return
        
        # ลบช่องว่างข้างหน้าหรือหลังข้อความในคอลัมน์ที่เป็นประเภท object (string)
        df = df.fillna('')
        df = df.astype(str)
        df = df.apply(lambda col: col.map(lambda x: x.strip() if isinstance(x, str) else x) if col.dtype == 'object' else col)
        df['PLANT_CODE'] = df['PLANT_CODE'].str.upper()
        df['SITE_ID'] = df['SITE_ID'].str.upper()
        
        # ตรวจสอบข้อมูลที่ซ้ำในไฟล์ Excel เองก่อน
        if df.duplicated(subset=['PLANT_CODE', 'SITE_ID']).any():
            duplicates = df[df.duplicated(subset=['PLANT_CODE', 'SITE_ID'], keep=False)]
            logger.error(f"Duplicate entries found in Excel file:\n{duplicates}")
            self.stdout.write(self.style.ERROR(f"Duplicate entries found in Excel file:\n{duplicates}"))
            return
        
        # จัดการการบันทึกข้อมูลภายใน transaction
        try:
            with transaction.atomic():
                # ลูปผ่านข้อมูลในแต่ละแถวและนำเข้าลงฐานข้อมูล
                for index, row in df.iterrows():
                    PLANT_CODE = row['PLANT_CODE']
                    SITE_ID = row['SITE_ID']

                    try:
                        # ดึง PlantType และ Site จากฐานข้อมูล
                        plant_type = PlantType.objects.get(plant_code=PLANT_CODE)
                        site = Site.objects.get(site_id=SITE_ID)
                        
                        # เพิ่มความสัมพันธ์ Many-to-Many
                        site.plant_types.add(plant_type)
                        site.save()

                        logger.info(f"Successfully added {SITE_ID} to {PLANT_CODE}")
                        self.stdout.write(self.style.SUCCESS(f"Successfully added {SITE_ID} to {PLANT_CODE}"))
                    
                    except PlantType.DoesNotExist:
                        logger.error(f"PlantType with code {PLANT_CODE} does not exist.")
                        self.stdout.write(self.style.ERROR(f"PlantType with code {PLANT_CODE} does not exist."))
                    except Site.DoesNotExist:
                        logger.error(f"Site with id {SITE_ID} does not exist.")
                        self.stdout.write(self.style.ERROR(f"Site with id {SITE_ID} does not exist."))
                
                logger.info("Data import completed successfully")
                self.stdout.write(self.style.SUCCESS("Data import completed successfully"))
        except Exception as e:
            logger.error(f"Transaction failed: {str(e)}")
            self.stdout.write(self.style.ERROR(f"Transaction failed: {str(e)}"))
