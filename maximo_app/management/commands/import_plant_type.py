import os
import pandas as pd
import logging
from django.conf import settings
from django.core.management.base import BaseCommand
from maximo_app.models import PlantType  # นำเข้าโมเดล PlantType
from django.db import transaction

# ตั้งค่า log สำหรับบันทึกข้อความการทำงาน
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class Command(BaseCommand):
    help = 'นำเข้าข้อมูลจากไฟล์ Excel สำหรับ PLANT_TYPE'

    def handle(self, *args, **kwargs):
        # เส้นทางไฟล์ Excel
        excel_file_path = os.path.join(settings.BASE_DIR, 'static', 'excel', 'Database.xlsx')

        # ตรวจสอบว่าไฟล์ Excel มีอยู่จริงหรือไม่
        if not os.path.exists(excel_file_path):
            logger.error(f"File {excel_file_path} not found")
            self.stdout.write(self.style.ERROR(f"File {excel_file_path} not found"))
            return

        try:
            # อ่านข้อมูลจาก Excel
            df = pd.read_excel(excel_file_path, sheet_name='PLANT_TYPE')
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
        required_columns = ['PLANT_CODE', 'PLANT_TYPE_TH', 'PLANT_TYPE_ENG']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            logger.error(f"Missing required columns in the Excel file: {', '.join(missing_columns)}")
            self.stdout.write(self.style.ERROR(f"Missing required columns in the Excel file: {', '.join(missing_columns)}"))
            return

        # แปลงข้อมูลให้เป็น format ที่เหมาะสม (ลบช่องว่างและแปลงเป็นตัวพิมพ์ใหญ่)
        df = df.fillna('')
        df = df.astype(str)
        df = df.apply(lambda col: col.map(lambda x: x.strip() if isinstance(x, str) else x) if col.dtype == 'object' else col)
        df['PLANT_CODE'] = df['PLANT_CODE'].str.upper()

        # ตรวจสอบข้อมูลที่ซ้ำในไฟล์ Excel เองก่อน
        if df['PLANT_CODE'].duplicated().any():
            duplicates = df[df['PLANT_CODE'].duplicated(keep=False)]
            logger.error(f"Duplicate entries found in Excel file:\n{duplicates}")
            self.stdout.write(self.style.ERROR(f"Duplicate entries found in Excel file:\n{duplicates}"))
            return

        # นำข้อมูลเข้าในฐานข้อมูล
        try:
            with transaction.atomic():  # ใช้ transaction เพื่อป้องกันการเปลี่ยนแปลงถ้ามีข้อผิดพลาด
                for index, row in df.iterrows():
                    # ตรวจสอบความยาวของฟิลด์
                    if len(row['PLANT_CODE']) > 8:
                        logger.warning(f"Skipping row {index+1}: PLANT_CODE exceeds maximum length (8 characters)")
                        self.stdout.write(self.style.WARNING(f"Skipping row {index+1}: PLANT_CODE exceeds maximum length (8 characters)"))
                        continue
                    if len(row['PLANT_TYPE_TH']) > 100:
                        logger.warning(f"Skipping row {index+1}: PLANT_TYPE_TH exceeds maximum length (100 characters)")
                        self.stdout.write(self.style.WARNING(f"Skipping row {index+1}: PLANT_TYPE_TH exceeds maximum length (100 characters)"))
                        continue
                    if len(row['PLANT_TYPE_ENG']) > 100:
                        logger.warning(f"Skipping row {index+1}: PLANT_TYPE_ENG exceeds maximum length (100 characters)")
                        self.stdout.write(self.style.WARNING(f"Skipping row {index+1}: PLANT_TYPE_ENG exceeds maximum length (100 characters)"))
                        continue

                    if pd.isna(row['PLANT_CODE']) or row['PLANT_CODE'] == '':
                        logger.warning(f"Skipping row {index+1}: Missing PLANT_CODE")
                        self.stdout.write(self.style.WARNING(f"Skipping row {index+1}: Missing PLANT_CODE"))
                        continue  # ข้ามแถวที่ข้อมูลไม่สมบูรณ์

                    # ใช้ get_or_create เพื่อสร้างข้อมูลใหม่ถ้าไม่มีข้อมูลในฐานข้อมูล
                    try:
                        plant_type, created = PlantType.objects.get_or_create(
                            plant_code=row['PLANT_CODE'],
                            defaults={
                                'plant_type_th': row['PLANT_TYPE_TH'],
                                'plant_type_eng': row['PLANT_TYPE_ENG'],
                            }
                        )
                        if created:
                            logger.info(f"Plant type {plant_type.plant_code} added")
                            self.stdout.write(self.style.SUCCESS(f"Plant type {plant_type.plant_code} added"))
                        else:
                            logger.warning(f"Plant type {plant_type.plant_code} already exists")
                            self.stdout.write(self.style.WARNING(f"Plant type {plant_type.plant_code} already exists"))

                    except Exception as e:
                        logger.error(f"Error processing row {index+1}: {str(e)}")
                        self.stdout.write(self.style.ERROR(f"Error processing row {index+1}: {str(e)}"))
                        continue  # ข้ามแถวที่เกิดข้อผิดพลาด

                logger.info("Data import completed successfully")
                self.stdout.write(self.style.SUCCESS("Data import completed successfully"))

        except Exception as e:
            logger.error(f"Transaction failed: {str(e)}")
            self.stdout.write(self.style.ERROR(f"Transaction failed: {str(e)}"))
