# path : maximo_project\maximo_app\forms.py
from datetime import datetime
from django import forms
from django.core.exceptions import ValidationError
from maximo_app.models import Site, ChildSite, PlantType, Unit, WorkType, ActType, WBSCode, Status

class UploadFileForm(forms.Form):
    schedule_file = forms.FileField(
        label='Schedule File', 
        widget=forms.ClearableFileInput(attrs={'accept': '.xlsx,.xlsm'})
    )
    location_file = forms.FileField(
        label='Location File', 
        widget=forms.ClearableFileInput(attrs={'accept': '.xlsx,.xlsm'})
    )
    
    # Dropdown สำหรับเลือกปี
    current_year = datetime.now().year
    YEAR_RANGE_CHOICES = [('', 'เลือก')] + [(str(year), str(year)) for year in range(current_year, current_year + 20)]
    
    year = forms.ChoiceField(
        label='YEAR', 
        choices=YEAR_RANGE_CHOICES, 
        required=True
    )
    
    FREQUENCY_CHOICES = [
        ('', 'เลือก'),
        ('1', '1 ปี'),
        ('2', '2 ปี'),
        ('3', '3 ปี'),
        ('4', '4 ปี'),
    ]
    
    frequency = forms.ChoiceField(
        label='FREQUENCY',
        choices=FREQUENCY_CHOICES,
        required=True,
    )
    
    plant_type = forms.ModelChoiceField(
        queryset=PlantType.objects.all(), 
        label='PLANT TYPE', 
        required=True, 
        empty_label="เลือก"
    )
    
    site = forms.ModelChoiceField(
        queryset=Site.objects.none(), 
        label='SITE', 
        required=True, 
        empty_label="เลือก"
    )
    
    child_site = forms.ModelChoiceField(
        queryset=ChildSite.objects.none(),
        label='CHILD SITE',
        required=True,
        empty_label="เลือก"
    )
    
    unit = forms.ModelChoiceField(
        queryset=Unit.objects.none(), 
        label='UNIT', 
        required=True, 
        empty_label="เลือก"
    )
    
    wostatus = forms.ModelChoiceField(
        queryset=Status.objects.all(), 
        label='STATUS', 
        required=True, 
        empty_label="เลือก"
    )
    
    work_type = forms.ModelChoiceField(
        queryset=WorkType.objects.none(), 
        label='WORKTYPE', 
        required=True, 
        empty_label="เลือก"
    )
    
    acttype = forms.ModelChoiceField(
        queryset=ActType.objects.none(), 
        label='ACTTYPE', 
        required=True, 
        empty_label="เลือก"
    )
    
    wbs = forms.ModelChoiceField(
        queryset=WBSCode.objects.all(), 
        label='SUBWBS GROUP', 
        required=True, 
        empty_label="เลือก"
    )

    def __init__(self, *args, **kwargs):
        super(UploadFileForm, self).__init__(*args, **kwargs)
        self.fields['plant_type'].queryset = PlantType.objects.all()
        self.fields['plant_type'].to_field_name = 'id'
        
        # กำหนด choices สำหรับ wbs และ wostatus
        self.fields['wbs'].choices = [
            ('', 'เลือก')
        ] + [
            (wbs.id, f'{wbs.wbs_code} - {wbs.description}')
            for wbs in WBSCode.objects.all()
        ]
        
        self.fields['wostatus'].choices = [
            ('', 'เลือก')
        ] + [
            (wo_status.id, f'{wo_status.status} - {wo_status.description}')
            for wo_status in Status.objects.all()
        ]
        
        # กรองข้อมูลจาก plant_type
        if 'plant_type' in self.data:
            try:
                plant_type_id = int(self.data.get('plant_type'))  # ดึง plant_type_id ที่ถูกเลือก
                plant_type = PlantType.objects.get(id=plant_type_id)  # ดึงข้อมูล PlantType ตาม plant_type_id

                # กรองข้อมูล acttype ที่เชื่อมโยงกับ plant_type
                self.fields['acttype'].queryset = plant_type.act_types.all()

                # กรองข้อมูล site ที่เชื่อมโยงกับ plant_type
                self.fields['site'].queryset = plant_type.sites.all()
                
                # กรองข้อมูล work_type ที่เชื่อมโยงกับ plant_type
                self.fields['work_type'].queryset = plant_type.work_types.all()
                
                # กรองข้อมูล unit ที่เชื่อมโยงกับ plant_type
                self.fields['unit'].queryset = plant_type.units.all()
                
            except (ValueError, TypeError, PlantType.DoesNotExist):
                # กำหนดให้ queryset เป็น none ถ้ามีข้อผิดพลาด
                self.fields['acttype'].queryset = ActType.objects.none()
                self.fields['site'].queryset = Site.objects.none()
                self.fields['work_type'].queryset = WorkType.objects.none()
                self.fields['unit'].queryset = Unit.objects.none()

        # กรองข้อมูลจาก site
        if 'site' in self.data:
            try:
                site_id = int(self.data.get('site'))  # ดึง site_id ที่ถูกเลือก
                site = Site.objects.get(id=site_id)  # ดึงข้อมูล Site ตาม site_id
                
                # กรองข้อมูล child_site ที่เชื่อมโยงกับ site
                self.fields['child_site'].queryset = site.child_sites.all()

            except (ValueError, TypeError, Site.DoesNotExist):
                # กำหนดให้ queryset เป็น none ถ้ามีข้อผิดพลาด
                self.fields['child_site'].queryset = ChildSite.objects.none()
    
    def clean(self):
        cleaned_data = super().clean()
        schedule_file = cleaned_data.get('schedule_file')
        location_file = cleaned_data.get('location_file')
        
        # ตรวจสอบประเภทของ schedule_file
        if schedule_file:
                content_type = schedule_file.content_type
                if content_type not in [
                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    'application/vnd.ms-excel.sheet.macroEnabled.12'
                ]:
                    self.add_error('schedule_file', 'เฉพาะไฟล์ .xlsx และ .xlsm เท่านั้น')
                # ฟังก์ชัน add_error จะเพิ่มข้อความข้อผิดพลาดให้กับฟิลด์ที่ไม่ผ่านการตรวจสอบ ข้อผิดพลาดนี้จะถูกแสดงในฟอร์มเมื่อผู้ใช้ทำการส่งฟอร์ม
        
        # ตรวจสอบประเภทของ location_file         
        if location_file:
                content_type = location_file.content_type
                if content_type not in [
                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    'application/vnd.ms-excel.sheet.macroEnabled.12'
                ]:
                    self.add_error('location_file', 'เฉพาะไฟล์ .xlsx และ .xlsm เท่านั้น')
                    
        # ตรวจสอบว่าฟิลด์ที่จำเป็นถูกเลือกหรือไม่
        required_fields = ['year', 'frequency', 'site', 'plant_type', 'unit', 'work_type', 'acttype', 'wbs', 'wostatus', 'child_site']
        for field in required_fields:
            if not cleaned_data.get(field):
                self.add_error(field, f'{field.upper()} ว่าง')

        return cleaned_data