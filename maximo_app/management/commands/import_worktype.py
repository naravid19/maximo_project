import os
import pandas as pd
import logging
from django.conf import settings
from django.core.management.base import BaseCommand
from maximo_app.models import WorkType  # นำเข้าโมเดล WorkType
from django.db import transaction

# ตั้งค่า log สำหรับบันทึกข้อความการทำงาน
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class Command(BaseCommand):
    help = 'นำเข้าข้อมูลจากไฟล์ Excel สำหรับ WORKTYPE'

    def handle(self, *args, **kwargs):
        # เส้นทางไฟล์ Excel
        excel_file_path = os.path.join(settings.BASE_DIR, 'static', 'excel', 'Database.xlsx')

        # ตรวจสอบว่าไฟล์ Excel มีอยู่จริงหรือไม่
        if not os.path.exists(excel_file_path):
            logger.error(f"File {excel_file_path} not found")
            self.stdout.write(self.style.ERROR(f"File {excel_file_path} not found"))
            return

        try:
            # อ่านข้อมูลจาก Excel โดยบังคับให้คอลัมน์ 'WORKTYPE' และ 'DESC' เป็น string
            df = pd.read_excel(excel_file_path, sheet_name='WORKTYPE', dtype={'WORKTYPE': str, 'DESC': str})
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

        # แปลงชื่อคอลัมน์เป็นตัวพิมพ์ใหญ่ทั้งหมด
        df.columns = df.columns.str.upper()

        # ตรวจสอบว่ามีคอลัมน์ 'WORKTYPE' และ 'DESC' หรือไม่
        required_columns = ['WORKTYPE', 'DESC']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            logger.error(f"Missing required columns in the Excel file: {', '.join(missing_columns)}")
            self.stdout.write(self.style.ERROR(f"Missing required columns in the Excel file: {', '.join(missing_columns)}"))
            return

        # ลบช่องว่างข้างหน้าหรือหลังข้อความและแทนที่ค่า NaN ด้วย ''
        df = df.fillna('')
        df = df.astype(str)
        df = df.apply(lambda col: col.map(lambda x: x.strip() if isinstance(x, str) else x) if col.dtype == 'object' else col)
        df['WORKTYPE'] = df['WORKTYPE'].str.upper()

        # ตรวจสอบข้อมูลที่ซ้ำในไฟล์ Excel เองก่อน
        if df['WORKTYPE'].duplicated().any():
            duplicates = df[df['WORKTYPE'].duplicated(keep=False)]
            logger.error(f"Duplicate entries found in Excel file:\n{duplicates}")
            self.stdout.write(self.style.ERROR(f"Duplicate entries found in Excel file:\n{duplicates}"))
            return

        # นำข้อมูลเข้าในฐานข้อมูล
        try:
            with transaction.atomic():
                for index, row in df.iterrows():
                    # ตรวจสอบความยาวของฟิลด์
                    if len(row['WORKTYPE']) > 8:
                        logger.warning(f"Skipping row {index+1}: WORKTYPE exceeds maximum length (8 characters)")
                        self.stdout.write(self.style.WARNING(f"Skipping row {index+1}: WORKTYPE exceeds maximum length (8 characters)"))
                        continue
                    if len(row['DESC']) > 100:
                        logger.warning(f"Skipping row {index+1}: DESC exceeds maximum length (100 characters)")
                        self.stdout.write(self.style.WARNING(f"Skipping row {index+1}: DESC exceeds maximum length (100 characters)"))
                        continue

                    if pd.isna(row['WORKTYPE']) or row['WORKTYPE'] == '':
                        logger.warning(f"Skipping row {index+1}: WORKTYPE is empty")
                        self.stdout.write(self.style.WARNING(f"Skipping row {index+1}: WORKTYPE is empty"))
                        continue

                    # ใช้ get_or_create เพื่อสร้างข้อมูลใหม่ถ้าไม่มีข้อมูลในฐานข้อมูล
                    try:
                        worktype, created = WorkType.objects.get_or_create(
                            worktype=row['WORKTYPE'],
                            defaults={
                                'description': row['DESC'],
                            }
                        )

                        if created:
                            logger.info(f"WorkType {worktype.worktype} added")
                            self.stdout.write(self.style.SUCCESS(f"WorkType {worktype.worktype} added"))
                        else:
                            logger.warning(f"WorkType {worktype.worktype} already exists")
                            self.stdout.write(self.style.WARNING(f"WorkType {worktype.worktype} already exists"))

                    except Exception as e:
                        logger.error(f"Error processing row {index+1}: {str(e)}")
                        self.stdout.write(self.style.ERROR(f"Error processing row {index+1}: {str(e)}"))
                        continue  # ข้ามแถวที่เกิดข้อผิดพลาด

                logger.info("Data import completed successfully")
                self.stdout.write(self.style.SUCCESS("Data import completed successfully"))

        except Exception as e:
            logger.error(f"Transaction failed: {str(e)}")
            self.stdout.write(self.style.ERROR(f"Transaction failed: {str(e)}"))
