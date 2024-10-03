import os
import pandas as pd
import logging
from django.conf import settings
from django.core.management.base import BaseCommand
from maximo_app.models import PlantType, WorkType
from django.db import transaction

# ตั้งค่า log สำหรับบันทึกข้อความการทำงาน
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class Command(BaseCommand):
    help = 'Import data from an Excel file for PlantType and WorkType'

    def handle(self, *args, **kwargs):
        # สร้างเส้นทางไปยังไฟล์ Database.xlsx ที่อยู่ใน static/excel
        excel_file_path = os.path.join(settings.BASE_DIR, 'static', 'excel', 'Database.xlsx')
        sheet_name = 'plant_worktype_relations'
        
        # ตรวจสอบว่าไฟล์ Excel มีอยู่หรือไม่
        if not os.path.exists(excel_file_path):
            logger.error(f"File {excel_file_path} not found")
            self.stdout.write(self.style.ERROR(f"File {excel_file_path} not found"))
            return
        
        try:
            # อ่านข้อมูลจาก Excel
            df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
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
        
        # แปลงชื่อคอลัมน์ให้เป็นตัวพิมพ์ใหญ่ทั้งหมด
        df.columns = df.columns.str.upper()
        
        # ตรวจสอบว่าคอลัมน์ที่จำเป็นมีอยู่ใน DataFrame หรือไม่
        required_columns = ['PLANT_CODE', 'WORKTYPE']
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
        df['WORKTYPE'] = df['WORKTYPE'].str.upper()
        # จัดการการบันทึกข้อมูลภายใน transaction
        try:
            with transaction.atomic():
                # ลูปผ่านข้อมูลในแต่ละแถวและนำเข้าลงฐานข้อมูล
                for index, row in df.iterrows():
                    plant_code = row['PLANT_CODE']
                    worktype_code = row['WORKTYPE']

                    # ตรวจสอบว่าฟิลด์มีค่าหรือไม่
                    if pd.isna(plant_code) or plant_code == '' or pd.isna(worktype_code) or worktype_code == '':
                        logger.warning(f"Skipping row {index+1}: Missing required fields")
                        self.stdout.write(self.style.WARNING(f"Skipping row {index+1}: Missing required fields"))
                        continue

                    try:
                        # ดึง PlantType และ WorkType จากฐานข้อมูล
                        plant = PlantType.objects.get(plant_code=plant_code)
                        worktype, created = WorkType.objects.get_or_create(worktype=worktype_code)

                        # เพิ่มความสัมพันธ์ Many-to-Many ระหว่าง PlantType และ WorkType
                        plant.work_types.add(worktype)

                        # บันทึกการเปลี่ยนแปลง
                        plant.save()

                        # แสดงข้อความสำเร็จถ้าข้อมูลถูกเพิ่มใหม่
                        logger.info(f"Successfully added worktype {worktype_code} to plant {plant_code}")
                        self.stdout.write(self.style.SUCCESS(f"Successfully added worktype {worktype_code} to plant {plant_code}"))

                    except PlantType.DoesNotExist:
                        self.stdout.write(self.style.ERROR(f"PlantType {plant_code} does not exist"))

                logger.info("Data import completed successfully")
                self.stdout.write(self.style.SUCCESS("Data import completed successfully"))

        except Exception as e:
            logger.error(f"Transaction failed: {str(e)}")
            self.stdout.write(self.style.ERROR(f"Transaction failed: {str(e)}"))
