import os
import pandas as pd
import logging
from django.conf import settings
from django.core.management.base import BaseCommand
from maximo_app.models import Status  # นำเข้าโมเดล Status
from django.db import transaction

# ตั้งค่า log สำหรับบันทึกข้อความการทำงาน
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class Command(BaseCommand):
    help = 'นำเข้าข้อมูลจากไฟล์ Excel สำหรับ Status'

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
            df = pd.read_excel(excel_file_path, sheet_name='STATUS')  # ใช้ sheet_name ให้ตรงกับในไฟล์ Excel
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
        required_columns = ['STATUS', 'DESC']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            logger.error(f"Missing required columns in the Excel file: {', '.join(missing_columns)}")
            self.stdout.write(self.style.ERROR(f"Missing required columns in the Excel file: {', '.join(missing_columns)}"))
            return

        # แปลงข้อมูลให้เป็น format ที่เหมาะสม
        df = df.fillna('')
        df = df.astype(str)
        df = df.apply(lambda col: col.map(lambda x: x.strip() if isinstance(x, str) else x) if col.dtype == 'object' else col)
        df['STATUS'] = df['STATUS'].str.upper()

        # ตรวจสอบข้อมูลที่ซ้ำในไฟล์ Excel เองก่อน
        if df['STATUS'].duplicated().any():
            duplicates = df[df['STATUS'].duplicated(keep=False)]
            logger.error(f"Duplicate entries found in Excel file:\n{duplicates}")
            self.stdout.write(self.style.ERROR(f"Duplicate entries found in Excel file:\n{duplicates}"))
            return

        # นำข้อมูลเข้าในฐานข้อมูล
        try:
            with transaction.atomic():
                for index, row in df.iterrows():
                    # ตรวจสอบความยาวของฟิลด์ต่างๆ
                    if len(row['STATUS']) > 8 or len(row['DESC']) > 100:
                        logger.warning(f"Skipping row {index+1}: Field length exceeded")
                        self.stdout.write(self.style.WARNING(f"Skipping row {index+1}: Field length exceeded"))
                        continue

                    # ข้ามแถวที่ข้อมูลสำคัญขาดหาย
                    if pd.isna(row['STATUS']) or row['STATUS'] == '' or pd.isna(row['DESC']) or row['DESC'] == '':
                        logger.warning(f"Skipping row {index+1}: Missing required fields")
                        self.stdout.write(self.style.WARNING(f"Skipping row {index+1}: Missing required fields"))
                        continue    # ข้ามแถวที่ข้อมูลไม่สมบูรณ์
                    
                    try:
                        status, created = Status.objects.get_or_create(
                            status=row['STATUS'],
                            defaults={
                                'description': row['DESC'],
                            }
                        )
                        if created:
                            logger.info(f"STATUS {status.status} added")
                            self.stdout.write(self.style.SUCCESS(f"STATUS {status.status} added"))
                        else:
                            logger.warning(f"STATUS {status.status} already exists")
                            self.stdout.write(self.style.WARNING(f"STATUS {status.status} already exists"))
                    
                    except Exception as e:
                        logger.error(f"Error processing row {index+1}: {str(e)}")
                        self.stdout.write(self.style.ERROR(f"Error processing row {index+1}: {str(e)}"))
                        continue  # ข้ามแถวที่เกิดข้อผิดพลาด
                        
                # แสดงข้อความเมื่อทำงานเสร็จสิ้น
                logger.info("Data import completed successfully")
                self.stdout.write(self.style.SUCCESS("Data import completed successfully"))
                
        except Exception as e:
            logger.error(f"Transaction failed: {str(e)}")
            self.stdout.write(self.style.ERROR(f"Transaction failed: {str(e)}"))
