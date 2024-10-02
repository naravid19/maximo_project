import os
import pandas as pd
import logging
from django.conf import settings
from django.core.management.base import BaseCommand
from maximo_app.models import PlantType, ActType
from django.db import transaction

# ตั้งค่า log สำหรับบันทึกข้อความการทำงาน
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class Command(BaseCommand):
    help = 'นำเข้าความสัมพันธ์ระหว่าง PlantType และ ActType จากไฟล์ Excel'

    def handle(self, *args, **kwargs):
        # สร้างเส้นทางไปยังไฟล์ Database.xlsx ที่อยู่ใน static/excel
        excel_file_path = os.path.join(settings.BASE_DIR, 'static', 'excel', 'Database.xlsx')
        
        # ตรวจสอบว่าไฟล์ Excel มีอยู่หรือไม่
        if not os.path.exists(excel_file_path):
            logger.error(f"File {excel_file_path} not found")
            self.stdout.write(self.style.ERROR(f"File {excel_file_path} not found"))
            return
        
        try:
            # อ่านข้อมูลจาก Excel
            df = pd.read_excel(excel_file_path, sheet_name='PlantType-ActType')
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
        required_columns = ['PLANT_CODE', 'ACTTYPE']
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
        df['ACTTYPE'] = df['ACTTYPE'].str.upper()
        
        # ตรวจสอบข้อมูลที่ซ้ำในไฟล์ Excel เองก่อน
        if df.duplicated(subset=['PLANT_CODE', 'ACTTYPE']).any():
            duplicates = df[df.duplicated(subset=['PLANT_CODE', 'ACTTYPE'], keep=False)]
            logger.error(f"Duplicate entries found in Excel file:\n{duplicates}")
            self.stdout.write(self.style.ERROR(f"Duplicate entries found in Excel file:\n{duplicates}"))
            return
        
        # จัดการการบันทึกข้อมูลภายใน transaction
        try:
            with transaction.atomic():
                # ลูปผ่านข้อมูลในแต่ละแถวและนำเข้าลงฐานข้อมูล
                for index, row in df.iterrows():
                    # ตรวจสอบความยาวของฟิลด์
                    if len(row['PLANT_CODE']) > 8:
                        logger.warning(f"Skipping row {index+1}: PLANT_CODE exceeds maximum length (8 characters)")
                        self.stdout.write(self.style.WARNING(f"Skipping row {index+1}: PLANT_CODE exceeds maximum length (8 characters)"))
                        continue
                    if len(row['ACTTYPE']) > 8:
                        logger.warning(f"Skipping row {index+1}: ACTTYPE exceeds maximum length (8 characters)")
                        self.stdout.write(self.style.WARNING(f"Skipping row {index+1}: ACTTYPE exceeds maximum length (8 characters)"))
                        continue

                    # ข้ามแถวที่ข้อมูลไม่สมบูรณ์
                    if pd.isna(row['PLANT_CODE']) or row['PLANT_CODE'] == '' or pd.isna(row['ACTTYPE']) or row['ACTTYPE'] == '':
                        logger.warning(f"Skipping row {index+1}: Missing required fields")
                        self.stdout.write(self.style.WARNING(f"Skipping row {index+1}: Missing required fields"))
                        continue
                    
                    try:
                        # ค้นหา PlantType และ ActType
                        plant_type = PlantType.objects.get(plant_code=row['PLANT_CODE'])
                        act_type = ActType.objects.get(acttype=row['ACTTYPE'])
                        
                        # เพิ่มความสัมพันธ์ Many-to-Many
                        plant_type.act_types.add(act_type)

                        # แสดงข้อความสำเร็จถ้าความสัมพันธ์ถูกเพิ่มใหม่
                        logger.info(f"Successfully linked PlantType {plant_type.plant_code} with ActType {act_type.acttype}")
                        self.stdout.write(self.style.SUCCESS(f"Successfully linked PlantType {plant_type.plant_code} with ActType {act_type.acttype}"))

                    except PlantType.DoesNotExist:
                        logger.error(f"PlantType with code {row['PLANT_CODE']} does not exist.")
                        self.stdout.write(self.style.ERROR(f"PlantType with code {row['PLANT_CODE']} does not exist."))
                    except ActType.DoesNotExist:
                        logger.error(f"ActType with code {row['ACTTYPE']} does not exist.")
                        self.stdout.write(self.style.ERROR(f"ActType with code {row['ACTTYPE']} does not exist."))

                logger.info("Data import completed successfully")
                self.stdout.write(self.style.SUCCESS("Data import completed successfully"))
        except Exception as e:
            logger.error(f"Transaction failed: {str(e)}")
            self.stdout.write(self.style.ERROR(f"Transaction failed: {str(e)}"))
