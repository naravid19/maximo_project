# path : maximo_project\maximo_app\forms.py
from datetime import datetime
from django import forms
from django.core.exceptions import ValidationError
from maximo_app.models import Site, ChildSite, PlantType, Unit, WorkType, ActType, WBSCode, Status
import os

GROUP_CHOICES = [
    ('no_arrange', 'No arrange group'),
    ('EGCRAFT', 'Craft'),
    ('PTW', 'Permit Type'),
    ('NEXTDATE', 'Start Date'),
    ('SYSTEM', 'System'),
]

class UploadFileForm(forms.Form):
    schedule_file = forms.FileField(
        label='Final Schedule File', 
        widget=forms.ClearableFileInput(attrs={'accept': '.xlsx,.xlsm'}),
        required=True
    )
    location_file = forms.FileField(
        label='Location File', 
        widget=forms.ClearableFileInput(attrs={'accept': '.xlsx,.xlsm'}),
        required=True
    )
    
    # Dropdown สำหรับเลือกปี
    current_year = datetime.now().year
    YEAR_RANGE_CHOICES = [('', 'เลือก')] + [(str(year), str(year)) for year in range(current_year, current_year + 10)]
    
    year = forms.ChoiceField(
        label='PLANT OUTAGE YEAR', 
        choices=YEAR_RANGE_CHOICES, 
        required=True
    )
    
    FREQUENCY_CHOICES = [
        ('', 'เลือก'),
        ('1', '1'),
        ('2', '2'),
        ('3', '3'),
        ('4', '4'),
        ('5', '5'),
        ('6', '6'),
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
        label='PLANT NAME',
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
        label='WORK TYPE', 
        required=True, 
        empty_label="เลือก"
    )
    
    acttype = forms.ModelChoiceField(
        queryset=ActType.objects.none(), 
        label='MNTACT TYPE', 
        required=True, 
        empty_label="เลือก"
    )
    
    wbs = forms.ModelChoiceField(
        queryset=WBSCode.objects.all(), 
        label='SUBWBS GROUP', 
        required=True, 
        empty_label="เลือก"
    )
    
    wbs_other = forms.CharField(
        label='Other WBS',
        required=False,
        widget=forms.TextInput(attrs={'id': 'other_wbs', 'disabled': 'true'}),
    )
    projectid_other = forms.CharField(
        label='Other PROJECTID',
        required=False,
        widget=forms.TextInput(attrs={'id': 'other_projectid', 'disabled': 'true'}),
    )
    
    selected_order = forms.CharField(required=True, widget=forms.HiddenInput())
    
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
    
    def clean_selected_order(self):
        selected_order = self.cleaned_data.get('selected_order', '')

        # แปลง selected_order ให้เป็น list หากเป็น string
        if isinstance(selected_order, str):
            selected_order_list = [item.strip() for item in selected_order.split(',') if item.strip()]
        else:
            selected_order_list = selected_order

        valid_choices = [choice[0] for choice in GROUP_CHOICES]
        for option in selected_order_list:
            if option not in valid_choices:
                self.add_error('selected_order', f'"{option}" ไม่ใช่ตัวเลือกที่ถูกต้อง')
        
        return selected_order_list

    def clean(self):
        cleaned_data = super().clean()
        wbs = cleaned_data.get('wbs')
        wbs_other = cleaned_data.get('wbs_other')
        projectid_other = cleaned_data.get('projectid_other')
        cleaned_data['selected_order'] = self.clean_selected_order()
        
        required_file_fields = ['schedule_file', 'location_file']
        for field in required_file_fields:
            if not cleaned_data.get(field):
                self.add_error(field, f'{field.upper()} ไม่มีไฟล์')
            
        allowed_mime_types = [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-excel.sheet.macroEnabled.12'
        ]
        
        allowed_extensions = ['.xlsx', '.xlsm']
        
        # ตรวจสอบประเภทของ schedule_file
        for file_field in ['schedule_file', 'location_file']:
            file = cleaned_data.get(file_field)
            if file:
                content_type = file.content_type
                ext = os.path.splitext(file.name)[1].lower()
                if content_type not in allowed_mime_types or ext not in allowed_extensions:
                    self.add_error(file_field, f'เฉพาะไฟล์ {", ".join(allowed_extensions)} เท่านั้น')
                    # ฟังก์ชัน add_error จะเพิ่มข้อความข้อผิดพลาดให้กับฟิลด์ที่ไม่ผ่านการตรวจสอบ ข้อผิดพลาดนี้จะถูกแสดงในฟอร์มเมื่อผู้ใช้ทำการส่งฟอร์ม
        
        # ตรวจสอบว่าฟิลด์ที่จำเป็นถูกเลือกหรือไม่
        required_fields = {
            'year': 'PLANT OUTAGE YEAR',
            'frequency': 'FREQUENCY',
            'plant_type': 'PLANT TYPE',
            'site': 'SITE',
            'child_site': 'PLANT NAME',
            'unit': 'UNIT',
            'wostatus': 'STATUS',
            'work_type': 'WORK TYPE',
            'acttype': 'MNTACT TYPE',
            'wbs': 'SUBWBS GROUP',
            
        }
        for field, field_label in required_fields.items():
            if not cleaned_data.get(field):
                self.add_error(field, f'กรุณาเลือก {field_label}')
        
        if wbs and wbs.wbs_code == 'อื่นๆ':
            if not wbs_other:
                self.add_error('wbs_other', 'กรุณากรอกข้อมูลในฟิลด์ "Other WBS" เนื่องจากคุณเลือก "อื่นๆ"')
            if not projectid_other:
                self.add_error('projectid_other', 'กรุณากรอกข้อมูลในฟิลด์ "Other PROJECTID" เนื่องจากคุณเลือก "อื่นๆ"')
        
        return cleaned_data