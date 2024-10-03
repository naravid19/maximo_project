import os
import pandas as pd
import logging
from django.conf import settings
from django.core.management.base import BaseCommand
from maximo_app.models import Unit
from django.db import transaction

# ตั้งค่า log สำหรับบันทึกข้อความการทำงาน
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class Command(BaseCommand):
    help = 'Import data from an Excel file for UNIT'

    def handle(self, *args, **kwargs):
        # เส้นทางไฟล์ Excel
        excel_file_path = os.path.join(settings.BASE_DIR, 'static', 'excel', 'Database.xlsx')
        sheet_name = 'unit'
        
        if not os.path.exists(excel_file_path):
            logger.error(f"File {excel_file_path} not found")
            self.stdout.write(self.style.ERROR(f"File {excel_file_path} not found"))
            return

        try:
            # อ่านข้อมูลจาก Excel โดยบังคับให้คอลัมน์ 'UNIT' เป็น string
            df = pd.read_excel(excel_file_path, sheet_name=sheet_name, dtype={'UNIT': str})
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
        
        df.columns = df.columns.str.upper()  # แปลงชื่อคอลัมน์เป็นตัวพิมพ์ใหญ่ทั้งหมด
        
        required_columns = ['UNIT', 'DESC']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            logger.error(f"Missing required columns in the Excel file: {', '.join(missing_columns)}")
            self.stdout.write(self.style.ERROR(f"Missing required columns in the Excel file: {', '.join(missing_columns)}"))
            return
        
        # ลบช่องว่างข้างหน้าและข้างหลัง และแทนที่ค่าว่าง/NaN ด้วย ''
        df = df.fillna('')
        df = df.astype(str)
        df = df.apply(lambda col: col.map(lambda x: x.strip() if isinstance(x, str) else x) if col.dtype == 'object' else col)

        # ตรวจสอบข้อมูลที่ซ้ำในไฟล์ Excel เองก่อน
        if df['UNIT'].duplicated().any():
            duplicates = df[df['UNIT'].duplicated(keep=False)]
            logger.error(f"Duplicate entries found in Excel file:\n{duplicates}")
            self.stdout.write(self.style.ERROR(f"Duplicate entries found in Excel file:\n{duplicates}"))
            return
        
        # ใช้ transaction.atomic เพื่อให้แน่ใจว่าข้อมูลทั้งหมดจะถูกบันทึกหรือย้อนกลับหากมีข้อผิดพลาด
        try:
            with transaction.atomic():
                for index, row in df.iterrows():
                    # ตรวจสอบความยาวของฟิลด์
                    if len(row['UNIT']) > 8:
                        logger.warning(f"Skipping row {index+1}: 'UNIT' exceeds maximum length (8 characters)")
                        self.stdout.write(self.style.WARNING(f"Skipping row {index+1}: 'UNIT' exceeds maximum length (8 characters)"))
                        continue
                    if len(row['DESC']) > 100:
                        logger.warning(f"Skipping row {index+1}: 'DESC' exceeds maximum length (100 characters)")
                        self.stdout.write(self.style.WARNING(f"Skipping row {index+1}: 'DESC' exceeds maximum length (100 characters)"))
                        continue

                    if pd.isna(row['UNIT']) or row['UNIT'] == '':
                        logger.warning(f"Skipping row {index+1}: 'UNIT' is empty or NaN")
                        self.stdout.write(self.style.WARNING(f"Skipping row {index+1}: 'UNIT' is empty or NaN"))
                        continue  # ข้ามแถวนี้ถ้า 'UNIT' ว่าง

                    unit, created = Unit.objects.get_or_create(
                        unit_code=row['UNIT'],
                        defaults={
                            'description': row['DESC'],
                        }
                    )
                    if created:
                        logger.info(f"Unit {unit.unit_code} added")
                        self.stdout.write(self.style.SUCCESS(f"Unit {unit.unit_code} added"))
                    else:
                        logger.warning(f"Unit {unit.unit_code} already exists")
                        self.stdout.write(self.style.WARNING(f"Unit {unit.unit_code} already exists"))

                logger.info("Data import completed successfully")
                self.stdout.write(self.style.SUCCESS("Data import completed successfully"))
        except Exception as e:
            logger.error(f"Transaction failed: {str(e)}")
            self.stdout.write(self.style.ERROR(f"Transaction failed: {str(e)}"))
