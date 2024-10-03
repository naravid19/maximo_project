import os
import pandas as pd
from django.conf import settings
from django.core.management.base import BaseCommand
from maximo_app.models import Site, ChildSite
from django.db import transaction

class Command(BaseCommand):
    help = 'Imports child site relations from an Excel file and links them to parent sites'

    def handle(self, *args, **kwargs):
        # สร้างเส้นทางไปยังไฟล์ Database.xlsx ที่อยู่ใน static/excel
        excel_file_path = os.path.join(settings.BASE_DIR, 'static', 'excel', 'Database.xlsx')
        sheet_name = 'child_site_relations'
        
        # ตรวจสอบว่าไฟล์ Excel มีอยู่หรือไม่
        if not os.path.exists(excel_file_path):
            self.stdout.write(self.style.ERROR(f"File {excel_file_path} not found"))
            return
        
        try:
            # อ่านข้อมูลจาก Excel
            df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
        except FileNotFoundError:
            self.stdout.write(self.style.ERROR("Error: The specified file was not found."))
            return
        except Exception as e:
            self.stdout.write(self.style.ERROR(f"Error reading the Excel file: {str(e)}"))
            return
        
        # แปลงชื่อคอลัมน์ให้เป็นตัวพิมพ์ใหญ่ทั้งหมด
        df.columns = df.columns.str.upper()

        # ตรวจสอบว่าคอลัมน์ที่จำเป็นมีอยู่ใน DataFrame หรือไม่
        required_columns = ['SITEID', 'CHILDID']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            self.stdout.write(self.style.ERROR(f"Missing required columns: {', '.join(missing_columns)}"))
            return
        
        # ใช้ transaction.atomic() เพื่อให้แน่ใจว่าการบันทึกข้อมูลจะเกิดขึ้นอย่างสมบูรณ์
        try:
            with transaction.atomic():
                for index, row in df.iterrows():
                    parent_site_id = row['SITEID'].strip().upper()
                    child_site_id = row['CHILDID'].strip().upper()

                    try:
                        # ดึง Site หลักจากฐานข้อมูล
                        parent_site = Site.objects.get(site_id=parent_site_id)
                        
                        # สร้างหรืออัพเดท ChildSite
                        child_site, created = ChildSite.objects.get_or_create(
                            site_id=child_site_id,
                            defaults={'parent_site': parent_site, 'site_name': f"Child {child_site_id}"}
                        )

                        # อัพเดท parent_site หาก ChildSite ถูกสร้างใหม่
                        if not created:
                            child_site.parent_site = parent_site
                            child_site.save()

                        self.stdout.write(self.style.SUCCESS(f"Successfully linked {child_site_id} to {parent_site_id}"))

                    except Site.DoesNotExist:
                        self.stdout.write(self.style.ERROR(f"Site with id {parent_site_id} does not exist."))

        except Exception as e:
            self.stdout.write(self.style.ERROR(f"Transaction failed: {str(e)}"))
