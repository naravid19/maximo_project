# path : maximo_project\maximo_app\forms.py
from datetime import datetime
from django import forms
from django.core.exceptions import ValidationError
from maximo_app.models import Site, PlantType, Unit, WorkType, ActType, WBSCode, Status

import mimetypes

class UploadFileForm(forms.Form):
    schedule_file = forms.FileField(label='Schedule File', widget=forms.ClearableFileInput(attrs={'accept': '.xlsx,.xlsm'}))
    location_file = forms.FileField(label='Location File', widget=forms.ClearableFileInput(attrs={'accept': '.xlsx,.xlsm'}))
    
    # Dropdown สำหรับเลือกปี
    current_year = datetime.now().year
    YEARS_CHOICES = [('', 'กรุณาเลือกปี')] + [(str(year), str(year)) for year in range(current_year, current_year + 20)]
    year = forms.ChoiceField(label='YEAR', choices=YEARS_CHOICES, required=True)
    
    # Dropdown สำหรับเลือกข้อมูล จากฐานข้อมูล
    site = forms.ModelChoiceField(queryset=Site.objects.all(), label='SITE', required=True, empty_label="กรุณาเลือก Site")
    plant_type = forms.ModelChoiceField(queryset=PlantType.objects.all(),label='PLANT TYPE', required=True, empty_label="กรุณาเลือก Plant Type")
    unit = forms.ModelChoiceField(queryset=Unit.objects.all(), label='UNIT', required=True, empty_label="กรุณาเลือก UNIT")
    work_type = forms.ModelChoiceField(queryset=WorkType.objects.filter(worktype__in=["APOO", "APAO"]), label='WORKTYPE', required=True, empty_label="กรุณาเลือก WORKTYPE")
    acttype = forms.ModelChoiceField(queryset=ActType.objects.none(), label='ACTTYPE', required=True, empty_label="กรุณาเลือก ACTTYPE")
    wbs = forms.ModelChoiceField(queryset=WBSCode.objects.all(), label='SUBWBS GROUP', required=True, empty_label="กรุณาเลือก Sub WBS Code")
    wostatus = forms.ModelChoiceField(queryset=Status.objects.all(), label='STATUS', required=True, empty_label="กรุณาเลือก WO Status")
    def __init__(self, *args, **kwargs):
        super(UploadFileForm, self).__init__(*args, **kwargs)
        self.fields['plant_type'].queryset = PlantType.objects.all()
        self.fields['plant_type'].to_field_name = 'plant_code'

        # ตรวจสอบว่ามีค่า plant_type ใน data หรือไม่
        plant_code = self.data.get('plant_type')
        if plant_code:
            try:
                # กรอง acttype ที่ remark ไม่เท่ากับ plant_code
                self.fields['acttype'].queryset = ActType.objects.exclude(remark=plant_code)
            except (ValueError, TypeError):
                self.fields['acttype'].queryset = ActType.objects.none()
        else:
            self.fields['acttype'].queryset = ActType.objects.none()
            
    def clean(self):
        cleaned_data = super().clean()
        schedule_file = cleaned_data.get('schedule_file')
        location_file = cleaned_data.get('location_file')
        
        # ตรวจสอบ mimetype ของ schedule_file
        if schedule_file:
            mimetype = mimetypes.guess_type(schedule_file.name)[0]
            if mimetype not in ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.ms-excel.sheet.macroEnabled.12']:
                self.add_error('schedule_file', 'เฉพาะไฟล์ .xlsx และ .xlsm เท่านั้น')
                # ฟังก์ชัน add_error จะเพิ่มข้อความข้อผิดพลาดให้กับฟิลด์ที่ไม่ผ่านการตรวจสอบ ข้อผิดพลาดนี้จะถูกแสดงในฟอร์มเมื่อผู้ใช้ทำการส่งฟอร์ม
        
        # ตรวจสอบ mimetype ของ location_file
        if location_file:
            mimetype = mimetypes.guess_type(location_file.name)[0]
            if mimetype not in ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.ms-excel.sheet.macroEnabled.12']:
                self.add_error('location_file', 'เฉพาะไฟล์ .xlsx และ .xlsm เท่านั้น')

        # ตรวจสอบว่าฟิลด์ที่จำเป็นถูกเลือกหรือไม่
        required_fields = ['year', 'site', 'plant_type', 'unit', 'work_type', 'acttype', 'wbs', 'wostatus']
        for field in required_fields:
            if not cleaned_data.get(field):
                self.add_error(field, f'กรุณาเลือก {field.upper()}')

        return cleaned_data