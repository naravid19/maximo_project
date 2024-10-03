# path : maximo_project/maximo_app/views.py

from .forms import UploadFileForm
from background_task import background
from django.conf import settings
from django.http import FileResponse, HttpResponse, Http404, JsonResponse
from django.shortcuts import render, HttpResponse, Http404
from django.views.decorators.http import require_GET
from django.urls import reverse
from maximo_app.models import Site, ChildSite, PlantType, Unit, WorkType, ActType, WBSCode, Status

import datetime
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, PatternFill, Alignment
from openpyxl.styles import Font
import logging
import os
import pandas as pd
import re
import shutil
import time
import uuid
import xlwings as xw

logger = logging.getLogger(__name__)

# Create your views here.

# ---------------------------------
# ฟังก์ชันพื้นฐาน
# ---------------------------------

def test(request):
    return render(request, 'maximo_app/usebase.html')

def index(request):
############
############
    schedule_filename = None
    location_filename = None
    extracted_kks_counts = None
    user_input_mapping = {}
    error_message = ""
############
############
    print(request.POST)
    print(request.FILES)
    print('0 def schedule_filename:',schedule_filename)
    print('0 def location_filename:',location_filename)
    print('0 def extracted_kks_counts:',extracted_kks_counts)
    print('0 def user_input_mapping:',user_input_mapping)
############
############  
    if request.method == 'POST':
        form = UploadFileForm(request.POST, request.FILES)
        if form.is_valid():
            ############
            ############
            print('1 form schedule_filename:',schedule_filename)
            print('1 form location_filename:',location_filename)
            print('1 form extracted_kks_counts:',extracted_kks_counts)
            print('1 form user_input_mapping:',user_input_mapping)
            ############
            ############
            # ตรวจสอบและสร้างโฟลเดอร์ temp
            temp_dir = 'temp'
            if not os.path.exists(temp_dir):
                os.makedirs(temp_dir)
            
            template_dir = os.path.join(settings.STATIC_ROOT, 'excel')
            if not os.path.exists(template_dir):
                os.makedirs(template_dir)
            
            schedule_file = request.FILES.get('schedule_file', None)
            location_file = request.FILES.get('location_file', None)
            child_site = form.cleaned_data.get('child_site')
            year = form.cleaned_data.get('year')
            site = form.cleaned_data.get('site')
            plant_type = form.cleaned_data.get('plant_type')
            unit = form.cleaned_data.get('unit')
            work_type = form.cleaned_data.get('work_type')
            acttype = form.cleaned_data.get('acttype')
            wbs = form.cleaned_data.get('wbs')
            wostatus = form.cleaned_data.get('wostatus')
            
            if year:
                logger.info(f"Year: {year}")
            if site:
                logger.info(f"Site: {site.site_id}")
            if plant_type:
                logger.info(f"Plant Type: {plant_type.plant_code}")
            if unit:
                logger.info(f"Unit: {unit.unit_code}")
            if work_type:
                logger.info(f"Work Type: {work_type.worktype}")
            if acttype:
                logger.info(f"ACTTYPE: {acttype.acttype}")
            if wbs:
                logger.info(f"WBS: {wbs.wbs_code}")
            if wostatus:
                logger.info(f"Status: {wostatus.status}")

            schedule_filename = schedule_file.name
            location_filename = location_file.name
            
            # สร้างชื่อไฟล์ที่ไม่ซ้ำกันโดยใช้ UUID
            unique_schedule_name = f"{uuid.uuid4()}_{schedule_file.name}"
            unique_location_name = f"{uuid.uuid4()}_{location_file.name}"
            
            schedule_path = os.path.join(temp_dir, unique_schedule_name)
            location_path = os.path.join(temp_dir, unique_location_name)
            comment_path = os.path.join(temp_dir, f"{uuid.uuid4()}_Comment.xlsx")
            
            logger.info(f"Files uploaded: Schedule file: {schedule_filename}, Location file: {schedule_filename}")
            
            try:
                with open(schedule_path, 'wb+') as destination:
                    for chunk in schedule_file.chunks():
                        destination.write(chunk)

                with open(location_path, 'wb+') as destination:
                    for chunk in location_file.chunks():
                        destination.write(chunk)
                        
                logger.info(f"Files saved successfully at: {schedule_path} and {location_path}")
            except IOError as e:
                logger.error(f"Error saving files: {str(e)}")
                return HttpResponse(f"Error saving files: {str(e)}")
############
############
            # Excel
            # sheet_name = 'Sheet1'
            
            orgid = 'EGAT'
            pluscrevum = 0
            status = 'ACTIVE'
            pluscjprevnum = 0
            frequency = 4
            frequnit = 'YEARS'
            route = ''
            leadtime = 7
            
            if acttype:
                egmntacttype = acttype.acttype
                
            if site:
                siteid = site.site_id   # 'SRD0'
                # site_str = re.sub(r'\d+', '', siteid)
                location = f'{child_site}-{plant_type}{unit}' # 'SRD-H02'
            
            if work_type:
                worktype = work_type.worktype   # 'APOO'
                
            
            if wostatus:
                wostatus = wostatus.status  # 'WSCH'
                
            if year:
                buddhist_year = int(year) + 543
                two_digits_year = buddhist_year % 100

            if acttype and wbs and two_digits_year:
                # plant_type = plant_type.plant_code
                # unit_code = unit.unit_code

                # สร้าง egprojectid และ egwbs
                egprojectid = f"O-{location.replace('-', '')}-{two_digits_year}{acttype.code}"  # 'O-SRDH02-67MI'
                egwbs = f"{egprojectid}-{wbs.wbs_code}" # 'O-SRDH02-67MI-WO'
                wbs_desc = f"{wbs.description} {acttype.description} {location} {buddhist_year}"

            logger.info(f"EGPROJECTID: {egprojectid}")
            logger.info(f"EGWBS: {egwbs}")
            logger.info(f"WBS DESC: {wbs_desc}")
            logger.info(f"Location: {location}") 
############
############
            # บันทึกค่าลงเซสชัน
            request.session['schedule_filename'] = schedule_file.name
            request.session['location_filename'] = location_file.name
            request.session['temp_dir'] = temp_dir
            
            request.session['egmntacttype'] = egmntacttype
            request.session['egprojectid'] = egprojectid
            request.session['egwbs'] = egwbs
            request.session['location'] = location
            request.session['siteid'] = siteid
            request.session['wbs_desc'] = wbs_desc
            request.session['worktype'] = worktype
            request.session['wostatus'] = wostatus

############
############
            try:
                required_columns = [
                    'KKS', 'EQUIPMENT', 'TASK_XX', 'TASK', 
                    'RESPONSE', 'DURATION_(HR.)', 'START_DATE', 'FINISH_DATE', 
                    'SUPERVISOR', 'FOREMAN', 'SKILL', 'RESPONSE_CRAFT', 
                    'ประเภทของ_PERMIT_TO_WORK', 'TYPE', 'COMMENT'
                ]
                
                important_columns = [
                    'KKS', 'EQUIPMENT', 'TASK_XX', 'TASK', 
                    'RESPONSE', 'DURATION_(HR.)', 'START_DATE', 'FINISH_DATE', 
                    'SUPERVISOR', 'FOREMAN', 'SKILL', 'RESPONSE_CRAFT', 
                    'ประเภทของ_PERMIT_TO_WORK', 'TYPE'
                ]

                df_original = read_excel_with_error_handling(schedule_path)
                if df_original is None:
                        raise ValueError("Cannot proceed without a valid DataFrame.")
                logger.info("Excel file loaded successfully.")
                
                # ลบช่องว่างที่ไม่จำเป็นออกจากชื่อคอลัมน์และแปลงชื่อคอลัมน์เป็นตัวพิมพ์ใหญ่ พร้อมแทนที่ช่องว่างด้วยขีดล่าง
                df_original.columns = df_original.columns.str.strip()
                df_original.columns = [col.upper().replace(' ', '_') for col in df_original.columns]
                logger.info("Column names cleaned and formatted.")
                
                # ตรวจสอบว่าคอลัมน์ที่จำเป็นทั้งหมดมีอยู่หรือไม่
                missing_columns = [col for col in required_columns if col not in df_original.columns]
                if missing_columns:
                    logger.error(f"Missing required columns: {', '.join(missing_columns)}")
                    raise ValueError(f"ข้อผิดพลาด: ไม่มีคอลัมน์ {', '.join(missing_columns)}")
                
                # ตรวจสอบว่าคอลัมน์ที่สำคัญมีข้อมูลอย่างน้อยหนึ่งค่า
                empty_columns = [col for col in important_columns if df_original[col].isna().all() or (df_original[col] == '').all()]
                if empty_columns:
                    logger.error(f"Important columns without data: {', '.join(empty_columns)}")
                    raise ValueError(f"ข้อผิดพลาด: คอลัมน์ต่อไปนี้ไม่มีข้อมูล: {', '.join(empty_columns)}")
                
                df_original = df_original[required_columns]
                df_original.rename(columns={'TASK_XX': 'TASK_ORDER'}, inplace=True)
                df_original['KKS'] = df_original['KKS'].str.upper()
                df_original['RESPONSE'] = df_original['RESPONSE'].str.upper()
                df_original['RESPONSE_CRAFT'] = df_original['RESPONSE_CRAFT'].str.upper()
                df_original['TYPE'] = df_original['TYPE'].str.upper()
                
                columns_to_strip = [
                    'KKS', 'EQUIPMENT', 'TASK', 'RESPONSE', 
                    'START_DATE', 'FINISH_DATE', 'RESPONSE_CRAFT', 
                    'ประเภทของ_PERMIT_TO_WORK', 'TYPE'
                ]
                for col in columns_to_strip:
                    if col in df_original.columns:
                        # ตรวจสอบว่าคอลัมน์เป็นชนิดข้อมูล string ก่อนใช้ .str.strip()
                        if df_original[col].dtype == "object":
                            df_original[col] = df_original[col].str.strip()
                            
                df_original_extracted = df_original
                
            except ValueError as ve:
                # ข้อผิดพลาดที่เกิดจากคอลัมน์ที่ขาดหายไป
                logger.error(f"Validation error: {str(ve)}")
                raise
            except Exception as e:
                logger.critical(f"Unexpected error: {str(e)}", exc_info=True)
                raise
            
            logger.info("DataFrame preparation completed successfully.")
            
            df_comment = pd.DataFrame()
            df_comment['KKS'] = df_original['KKS']
            df_comment['COMMENT'] = ''

            df_original = df_original.replace('', np.nan)

            long_equipment = df_original['EQUIPMENT'].str.len()>100
            update_comment(df_comment, long_equipment, 'COMMENT', 'EQUIPMENT มีความยาวมากกว่า 100 ตัวอักษร')

            # กรองข้อมูลตาม task_with_order และความยาวของสตริงใน TASK
            task_with_order = ((df_original['TASK_ORDER'].notna())) & ((df_original['TASK_ORDER']!='xx'))
            long_task = df_original['TASK'].str.len() > 100
            update_comment(df_comment, long_task, 'COMMENT', 'TASK มีความยาวมากกว่า 100 ตัวอักษร')

            task_na = df_original['TASK'].isna()
            update_comment(df_comment, (task_with_order & task_na), 'COMMENT', 'ไม่มี TASK')

            # ใช้ฟังก์ชัน convert_duration กับคอลัมน์ DURATION_(HR.)
            df_original['DURATION_(HR.)'] = df_original['DURATION_(HR.)'].apply(convert_duration)

            lst_index_duration = []
            for i,j in zip(df_original['DURATION_(HR.)'], df_original.index) :
                #if (type(i)!=float) and (type(i)!=int):
                if (type(i)==str): ## if DURATION_(HR.) is a string type then show value and index
                    i,j
                    lst_index_duration.append(j)
            print('##############################')
            lst_index_start = []
            for i,j in zip(df_original['START_DATE'], df_original.index) :
                if (type(i)==str):
                    if (re.search("[a-zA-Z]",i)):## if found alphabt in START_DATE then show index
                        #i,j
                        lst_index_start.append(j)
            print('##############################')
            lst_index_finish = []
            for i,j in zip(df_original['FINISH_DATE'], df_original.index):
                if (type(i)==str):
                    if (re.search("[a-zA-Z]",i)):## if found alphabt in FINISH_DATE then show index
                        #i,j
                        lst_index_finish.append(j)
            
            # ตรวจสอบค่า DURATION_(HR.) ว่าเป็นตัวเลขหรือไม่
            cond_duration = df_original['DURATION_(HR.)'].apply(lambda x: isinstance(x, (int, float)))
            non_negative = df_original[cond_duration]['DURATION_(HR.)'] < 0
            task_order_not_xx = df_original['TASK_ORDER'] != 'xx'
            update_comment(df_comment, ((~cond_duration | non_negative) & task_order_not_xx), 'COMMENT', 'DURATION_(HR.) ไม่ถูกต้อง')

            cond_start_date = df_original['START_DATE'].apply(lambda x: isinstance(x, str) and re.search("[a-zA-Z]", x) is not None and not is_date(x))
            update_comment(df_comment, cond_start_date, 'COMMENT', 'START_DATE มีตัวอักษร')

            cond_finish_date = df_original['FINISH_DATE'].apply(lambda x: isinstance(x, str) and re.search("[a-zA-Z]", x) is not None and not is_date(x))
            update_comment(df_comment, cond_finish_date, 'COMMENT', 'FINISH_DATE มีตัวอักษร')
            
            
            # df_original_strat_date_1 = pd.to_datetime(df_original['START_DATE'], format='%d/%m/%Y',errors='coerce')
            # df_original_strat_date_2 = pd.to_datetime(df_original['START_DATE'], format='%Y-%m-%d %H:%M:%S',errors='coerce')
            # df_original_strat_date_3 = pd.to_datetime(df_original['START_DATE'], format='%d-%b-%Y', errors='coerce')
            # df_original_strat_date_3 = df_original_strat_date_3.dt.strftime('%Y-%m-%d %H:%M:%S')
            # df_original['START_DATE_NEW'] = df_original_strat_date_1.fillna(df_original_strat_date_2)
            # df_original['START_DATE_NEW'] = df_original['START_DATE_NEW'].fillna(df_original_strat_date_3)

            # df_original_finish_date_1 = pd.to_datetime(df_original['FINISH_DATE'], format='%d/%m/%Y',errors='coerce')
            # df_original_finish_date_2 = pd.to_datetime(df_original['FINISH_DATE'], format='%Y-%m-%d %H:%M:%S',errors='coerce')
            # df_original_finish_date_3 = pd.to_datetime(df_original['FINISH_DATE'], format='%d-%b-%Y', errors='coerce')
            # df_original_finish_date_3 = df_original_finish_date_3.dt.strftime('%Y-%m-%d %H:%M:%S')
            # df_original['FINISH_DATE_NEW'] = df_original_finish_date_1.fillna(df_original_finish_date_2)
            # df_original['FINISH_DATE_NEW'] = df_original['FINISH_DATE_NEW'].fillna(df_original_finish_date_3)

            # ใช้ฟังก์ชัน parse_dates กับ START_DATE
            df_original['START_DATE_NEW'] = parse_dates(df_original['START_DATE'])

            # ใช้ฟังก์ชัน parse_dates กับ FINISH_DATE
            df_original['FINISH_DATE_NEW'] = parse_dates(df_original['FINISH_DATE'])

            cond_start_date_new = df_original['START_DATE'].notna() & df_original['START_DATE_NEW'].isna() & (df_original['TASK_ORDER'].notna()) & (df_original['TASK_ORDER'] != 'xx')
            update_comment(df_comment, cond_start_date_new, 'COMMENT', 'START_DATE ไม่ถูกต้อง')

            cond_finish_date_new = df_original['FINISH_DATE'].notna() & df_original['FINISH_DATE_NEW'].isna() & (df_original['TASK_ORDER'].notna()) & (df_original['TASK_ORDER'] != 'xx')
            update_comment(df_comment, cond_finish_date_new, 'COMMENT', 'FINISH_DATE ไม่ถูกต้อง')

            kks_length_exceed = df_original['KKS'].str.len() > 30 
            update_comment(df_comment, kks_length_exceed, 'COMMENT', 'KKS มีความยาวมากกว่า 30 ตัวอักษร')

            # หาค่า Plant Unit ที่พบมากที่สุด
            most_common_plant_unit = df_original['KKS'].str.split('-', expand=True)[0].value_counts().idxmax()
            # สร้างเงื่อนไขสำหรับค่า Plant Unit ที่ไม่ตรงกับค่าที่พบมากที่สุด
            cond_kks = (df_original['KKS'].str.split('-', expand=True)[0] != most_common_plant_unit) & (df_original['KKS'].notna())
            update_comment(df_comment, cond_kks, 'COMMENT', 'Plant Unit ไม่สอดคล้อง')

            plant_list = df_original['KKS'].str.split('-',expand =True)[0].value_counts().index.tolist()
            plant_regex = "|".join([p + "-" for p in plant_list])
            first_plant = plant_regex.split('|')[0]

            #df_original['KKS'] = df_original['KKS'].replace(r'^TTN-','', regex=True)
            df_original['KKS'] = df_original['KKS'].replace(plant_regex,'', regex=True)

            # cond1 = (df_original['TASK_ORDER']==10) & ((df_original['DURATION (HR.)'].isna())|(df_original['START DATE_new'].isna())|(df_original['FINISH DATE_new'].isna()))
            # cond2 = (df_original['TASK_ORDER']==10) & ((df_original['SUPERVISOR'].isna())&(df_original['FOREMAN'].isna())&(df_original['SKILL'].isna())&(df_original['RESPONSE CRAFT'].isna()))
            # cond3 = (df_original['TASK_ORDER']==10) & (df_original['SUPERVISOR'].isna())&(df_original['FOREMAN'].isna())&(df_original['SKILL'].isna())
            task_no_start_date = (df_original['TASK_ORDER'].notna()) & (df_original['TASK_ORDER'] != 'xx') & (df_original['START_DATE_NEW'].isna())
            task_no_finish_date = (df_original['TASK_ORDER'].notna()) & (df_original['TASK_ORDER'] != 'xx') & (df_original['FINISH_DATE_NEW'].isna())
            task_no_skill_rate = (df_original['TASK_ORDER'] == 10) & ((df_original['SUPERVISOR'].isna()) & (df_original['FOREMAN'].isna()) & (df_original['SKILL'].isna()))
            task_no_craft = (df_original['TASK_ORDER'].notna()) & (df_original['TASK_ORDER'] != 'xx') & (df_original['RESPONSE_CRAFT'].isna())
            task_no_response = (df_original['TASK_ORDER'].notna()) & (df_original['TASK_ORDER'] != 'xx') & (df_original['RESPONSE'].isna())

            # df_original.loc[cond1, 'COMMENT'] = 'ไม่มีวันที่'
            update_comment(df_comment, task_no_start_date, 'COMMENT', 'ไม่มี START_DATE')

            # df_original.loc[cond1, 'COMMENT'] = 'ไม่มีวันที่'
            update_comment(df_comment, task_no_finish_date, 'COMMENT', 'ไม่มี FINISH_DATE')

            # df_original.loc[cond3, 'COMMENT'= 'ไม่มี skill rate'
            update_comment(df_comment, task_no_skill_rate, 'COMMENT', 'ไม่มี SKILL RATE (ต้องมี)')


            task_order = (df_original['TASK_ORDER'].notna()) & (df_original['TASK_ORDER'] != 'xx')
            # ตรวจสอบว่า SUPERVISOR, FOREMAN, SKILL ถ้ามีค่า (ไม่ใช่ NaN) ต้องเป็นจำนวนเต็มบวก
            valid_supervisor = df_original['SUPERVISOR'].apply(lambda x: (pd.isna(x) or (isinstance(x, (int, float)) and x.is_integer() and x >= 0)))
            valid_foreman = df_original['FOREMAN'].apply(lambda x: (pd.isna(x) or (isinstance(x, (int, float)) and x.is_integer() and x >= 0)))
            valid_skill = df_original['SKILL'].apply(lambda x: (pd.isna(x) or (isinstance(x, (int, float)) and x.is_integer() and x >= 0)))
            # เงื่อนไขตรวจสอบว่าค่าใน SUPERVISOR, FOREMAN, SKILL มีค่าที่ไม่ถูกต้อง (ไม่เป็นจำนวนเต็มบวก)
            invalid_skill_rate = ~(valid_supervisor & valid_foreman & valid_skill)
            update_comment(df_comment, (task_order & invalid_skill_rate), 'COMMENT', 'SKILL RATE ไม่ถูกต้อง')

            # df_original.loc[cond2, 'COMMENT'] = 'ไม่มี RESPONSE_CRAFT'
            update_comment(df_comment, task_no_craft, 'COMMENT', 'ไม่มี RESPONSE_CRAFT')

            # RESPONSE_CRAFT มีความยาวเกิน 12 อักขระ
            carft_length_exceed = df_original['RESPONSE_CRAFT'].str.len() > 12 
            update_comment(df_comment, carft_length_exceed, 'COMMENT', 'RESPONSE_CRAFT มีความยาวมากกว่า 12 ตัวอักษร')

            # ไม่มี RESPONSE
            update_comment(df_comment, task_no_response, 'COMMENT', 'ไม่มี RESPONSE')

            # RESPONSE มีความยาวเกิน 12 อักขระ
            response_length_exceed = df_original['RESPONSE'].str.len() > 12 
            update_comment(df_comment, response_length_exceed, 'COMMENT', 'RESPONSE มีความยาวมากกว่า 12 ตัวอักษร')

            df_original_copy = df_original.copy()
            df_original_copy = df_original_copy.dropna(how='all')
            df_original_copy = df_original_copy.reset_index(drop=True)
            df_original_copy = df_original_copy.fillna(-1)

            temp_kks = ''##kks
            lst_kks = []
            ############
            temp_eq = ''##equipment
            lst_eq = []
            ###############
            temp_task_order = ''##Task_order
            lst_task_order = []
            ################
            temp_task = ''##Task
            lst_task = []
            #################
            temp_start_date = ''##start_date
            lst_start_date = []
            ##################
            temp_finish_date = ''##finish_date
            lst_finish_date = []
            ##################
            temp_type_work = ''##PTW
            lst_type_work = []
            #####################
            temp_response = ''##response
            lst_response = []
            #####################
            temp_RESPONSE_CRAFT = ''##response carft
            lst_RESPONSE_CRAFT = []
            #####################
            for i in (df_original_copy['KKS']):
                if i!=-1:
                    temp_kks = i
                    lst_kks.append(i)
                else:
                    i=temp_kks
                    lst_kks.append(i)
            ######################
            for j in (df_original_copy['EQUIPMENT']):
                if j!=-1:
                    temp_eq = j
                    lst_eq.append(j)
                else:
                    j=temp_eq
                    lst_eq.append(j)
            #######################
            for k in (df_original_copy['TASK_ORDER']):
                if k!=-1:
                    temp_task_order = k
                    lst_task_order.append(k)
                else:
                    k=temp_task_order
                    lst_task_order.append(k)
            ######################
            for z in (df_original_copy['TASK']):
                if z!=-1:
                    temp_task = z
                    lst_task.append(z)
                else:
                    z=temp_task
                    lst_task.append(z)
            ######################
            for xx in (df_original_copy['START_DATE_NEW']):
                if xx!=-1:
                    temp_start_date = xx
                    lst_start_date.append(xx)
                else:
                    xx=temp_start_date
                    lst_start_date.append(xx)
            ######################
            for yy in (df_original_copy['FINISH_DATE_NEW']):
                if yy!=-1:
                    temp_finish_date = yy
                    lst_finish_date.append(yy)
                else:
                    yy=temp_finish_date
                    lst_finish_date.append(yy)
            ######################
            for zz in (df_original_copy['ประเภทของ_PERMIT_TO_WORK']):
                if zz!=-1:
                    temp_type_work = zz
                    lst_type_work.append(zz)
                else:
                    zz=temp_type_work
                    lst_type_work.append(zz)
            ######################
            for xy in (df_original_copy['RESPONSE']):
                if xy!=-1:
                    temp_response = xy
                    lst_response.append(xy)
                else:
                    xy=temp_response
                    lst_response.append(xy)
            ######################
            for xy in (df_original_copy['RESPONSE_CRAFT']):
                if xy!=-1:
                    temp_RESPONSE_CRAFT = xy
                    lst_RESPONSE_CRAFT.append(xy)
                else:
                    xy=temp_RESPONSE_CRAFT
                    lst_RESPONSE_CRAFT.append(xy)

            df_original_copy['KKS_NEW'] = pd.Series(lst_kks)
            df_original_copy['EQUIPMENT_NEW'] = pd.Series(lst_eq)
            df_original_copy['TASK_ORDER_NEW'] = pd.Series(lst_task_order)
            df_original_copy['TASK_NEW']= pd.Series(lst_task)
            ###################
            ###################
            df_original_copy['START_DATE']=pd.Series(lst_start_date)
            df_original_copy['FINISH_DATE']=pd.Series(lst_finish_date)
            df_original_copy['ประเภทของ_PERMIT_TO_WORK'] = pd.Series(lst_type_work)
            df_original_copy['RESPONSE'] = pd.Series(lst_response)
            df_original_copy['RESPONSE_CRAFT'] = pd.Series(lst_RESPONSE_CRAFT)

            lst = ['KKS_NEW', 'EQUIPMENT_NEW', 'TASK_ORDER_NEW', 'TASK_NEW', 'RESPONSE',
                'DURATION_(HR.)', 'START_DATE', 'FINISH_DATE', 'SUPERVISOR',
                'FOREMAN', 'SKILL', 'RESPONSE_CRAFT', 'ประเภทของ_PERMIT_TO_WORK','TYPE']
            df_original_newcol = df_original_copy.loc[:,lst].copy()
            main_system = df_original_newcol['KKS_NEW'].str[0:6]
            sub_system =  df_original_newcol['KKS_NEW'].str[0:8]
            equipment =   df_original_newcol['KKS_NEW'].str[8:10]

            df_original_newcol['MAIN_SYSTEM'] = main_system
            df_original_newcol['SUB_SYSTEM'] = sub_system
            df_original_newcol['EQUIPMENT'] = equipment

            import warnings
            warnings.simplefilter("ignore")
            kks_read = pd.read_excel(location_path, 
                                sheet_name=0,header = 0,
                                usecols = 'A:B',
                                engine="openpyxl")

            plant_list1 = kks_read['Location'].str.split('-',expand=True)[0].value_counts().index.tolist()
            plant_regex1 = "|".join([p + "-" for p in plant_list1])
            kks_read['Location_x'] = kks_read['Location'].str.replace(plant_regex1, '', regex=True)

            lst_main_sys = []
            for i,j in df_original_newcol.iterrows():
                if len(kks_read[kks_read['Location_x']==j['MAIN_SYSTEM']]['Description'])==1:
                    desc = kks_read[kks_read['Location_x']==j['MAIN_SYSTEM']]['Description'].values[0]
                    lst_main_sys.append(desc)
                else:
                    lst_main_sys.append('No_kks_found')

            lst_sub_sys = []
            for i,j in df_original_newcol.iterrows():
                if len(kks_read[kks_read['Location_x']==j['SUB_SYSTEM']]['Description'])==1:
                    desc = kks_read[kks_read['Location_x']==j['SUB_SYSTEM']]['Description'].values[0]
                    lst_sub_sys.append(desc)
                else:
                    lst_sub_sys.append('No_kks_found')


            lst_equipment = []
            for i,j in df_original_newcol.iterrows():
                if len(kks_read[kks_read['Location_x']==j['KKS_NEW']]['Description'])==1:
                    desc = kks_read[kks_read['Location_x']==j['KKS_NEW']]['Description'].values[0]
                    lst_equipment.append(desc)
                else:
                    lst_equipment.append('No_kks_found')

            main_sys = pd.Series(lst_main_sys)
            sub_sys = pd.Series(lst_sub_sys)
            eq_sys = pd.Series(lst_equipment)
            df_original_newcol['MAIN_SYSTEM_DESC']=main_sys
            df_original_newcol['SUB_SYSTEM_DESC']=sub_sys
            df_original_newcol['KKS_NEW_DESC']=eq_sys

            cond1 = df_original_newcol['KKS_NEW'].notna()
            cond2 = df_original_newcol['MAIN_SYSTEM_DESC']=='No_kks_found'
            cond3 = df_original_newcol['SUB_SYSTEM_DESC']=='No_kks_found'
            cond4 = df_original_newcol['KKS_NEW_DESC']=='No_kks_found'
            cond5 = df_original_newcol['KKS_NEW'].isin([''])
            cond6 = df_original_newcol['TASK_ORDER_NEW']!='xx'

            update_comment(df_comment, (cond1 & cond4 & ~cond5 & ((df_original_copy['KKS']!= -1))), 'COMMENT', 'ไม่พบ kks')

            # DURATION_(HR.)
            task_no_duration = (df_original['TASK_ORDER'].notna()) & (df_original['TASK_ORDER'] != 'xx') & (df_original['DURATION_(HR.)'].isna())
            update_comment(df_comment, task_no_duration, 'COMMENT', 'ไม่มี DURATION_(HR.)')

            # DURATION_(HR.) ยาวเกิน 8 อักขระ
            # กรองเฉพาะค่าที่ไม่มี NaN และมีค่ามากกว่า 8 หลัก
            duration_length_exceed = df_original['DURATION_(HR.)'].apply(
                lambda x: len(f'{int(x):.0f}') > 8 if pd.notna(x) and isinstance(x, (int, float)) else False
            )
            update_comment(df_comment, duration_length_exceed, 'COMMENT', 'DURATION_(HR.) มีความยาวมากกว่า 8 หลัก')

            # ประเภทของ_PERMIT_TO_WORK
            task_no_ptw = (df_original['TASK_ORDER'].notna()) & (df_original['TASK_ORDER'] != 'xx') & (df_original['ประเภทของ_PERMIT_TO_WORK'].isna())
            update_comment(df_comment, task_no_ptw, 'COMMENT', 'ไม่มี ประเภทของ_PERMIT_TO_WORK')

            # ประเภทของ_PERMIT_TO_WORK ความยาวเกิน 250 อักขระ
            ptw_length_exceed = df_original['ประเภทของ_PERMIT_TO_WORK'].str.len() > 250
            update_comment(df_comment, ptw_length_exceed, 'COMMENT', 'ประเภทของ_PERMIT_TO_WORK มีความยาวมากกว่า 250 ตัวอักษร')

            # TYPE
            task_no_type = (df_original['TASK_ORDER'].notna()) & (df_original['TASK_ORDER'] != 'xx') & (df_original['TYPE'].isna())
            update_comment(df_comment, task_no_type, 'COMMENT', 'ไม่มี TYPE (ต้องมี)')

            # TYPE ไม่ถูกต้อง
            task_type = (df_original['TASK_ORDER'].notna()) & (df_original['TASK_ORDER'] != 'xx') & (df_original['TYPE'].notna())
            valid_type = (df_original['TYPE'].isin(['ME', 'EE', 'CV', 'IC']))
            update_comment(df_comment, (task_type & ~valid_type), 'COMMENT', 'TYPE ไม่ถูกต้อง')

            # TASK_ORDER
            # TASK_ORDER ไม่ถูกต้อง
            df_original_check = df_original_copy.copy()
            df_original_check['TASK_ORDER_NEW'] = df_original_check['TASK_ORDER_NEW'].astype(str)
            # ตรวจสอบว่าค่าใน TASK_ORDER_NEW เป็นจำนวนเต็มบวกหรือ 'xx' เท่านั้น
            task_order_valid = df_original_check['TASK_ORDER_NEW'].apply(
                lambda x: (isinstance(x, str) and (x.isdigit() or x == 'xx')) or 
                        (isinstance(x, (int, float)) and x.is_integer() and x >= 0)
            )

            update_comment(df_comment, (~task_order_valid & (df_original_check['TASK_ORDER']!= -1)), 'COMMENT', 'TASK_ORDER ไม่ถูกต้อง')

            # TASK_ORDER ไม่มี
            duration_notna_cond = df_original['DURATION_(HR.)'].notna()
            task_missing_cond = df_original['TASK_ORDER'].isna() | (df_original['TASK_ORDER'] == '')
            update_comment(df_comment, (duration_notna_cond & task_missing_cond), 'COMMENT', 'ไม่มี TASK_ORDER')

            # TASK_ORDER ยาวเกิน 12 อักขระ
            task_order_length_exceed = df_original_check['TASK_ORDER_NEW'].str.len() > 12
            update_comment(df_comment, (task_order_length_exceed & (df_original_check['TASK_ORDER']!= -1)), 'COMMENT', 'TASK_ORDER มีความยาวมากกว่า 12 ตัวอักษร')


            kks_new_na = df_original_copy['KKS_NEW'].isin([''])
            equip_new_na = df_original_copy['EQUIPMENT_NEW'].isin([''])
            task_order_new_na = df_original_copy['TASK_ORDER_NEW'].isin([''])
            task_new_na = df_original_copy['TASK_NEW'].isin([''])
            start_date_na = df_original_copy['START_DATE'].isin([''])
            finish_date_na = df_original_copy['FINISH_DATE'].isin([''])
            ptw_na = df_original_copy['ประเภทของ_PERMIT_TO_WORK'].isin([''])
            response_new_na = df_original_copy['RESPONSE'].isin([''])
            response_craft_na = df_original_copy['RESPONSE_CRAFT'].isin([''])

            task_not_xx = df_original_copy['TASK_ORDER'] != 'xx'
            # comment_empty_or_na = (df_comment['COMMENT'] == '') | (df_comment['COMMENT'].isna())
            # df_comment.loc[task_not_xx & kks_new_na & comment_empty_or_na, 'COMMENT'] = 'ไม่มี KKS (ต้องมี)'
            # # df_comment.loc[task_not_xx & kks_new_na & ~comment_empty_or_na, 'COMMENT'] += ', ไม่มี KKS'
            # df_comment.loc[task_not_xx & kks_new_na & ~comment_empty_or_na, 'COMMENT'] = df_comment.loc[task_not_xx & kks_new_na & ~comment_empty_or_na, 'COMMENT'].apply(
            #     lambda x: x.replace('ไม่มี KKS', 'ไม่มี KKS (ต้องมี)') if 'ไม่มี KKS' in x else f"{x}, ไม่มี KKS (ต้องมี)")
            replace_or_append_comment(df_comment, (task_not_xx & kks_new_na), 'COMMENT', 'ไม่มี KKS (ต้องมี)', replace_message='ไม่มี KKS')

            # df_comment.loc[task_not_xx & equip_new_na & comment_empty_or_na, 'COMMENT'] = 'ไม่มี EQUIPMENT (ต้องมี)'
            # df_comment.loc[task_not_xx & equip_new_na & ~comment_empty_or_na, 'COMMENT'] += ', ไม่มี EQUIPMENT (ต้องมี)'
            update_comment(df_comment, (task_not_xx & equip_new_na), 'COMMENT', 'ไม่มี EQUIPMENT (ต้องมี)')


            # df_comment.loc[task_not_xx & task_order_new_na & comment_empty_or_na, 'COMMENT'] = 'ไม่มี TASK_ORDER (ต้องมี)'
            # df_comment.loc[task_not_xx & task_order_new_na & ~comment_empty_or_na, 'COMMENT'] += ', ไม่มี TASK_ORDER (ต้องมี)'
            update_comment(df_comment, (task_not_xx & task_order_new_na), 'COMMENT', 'ไม่มี TASK_ORDER (ต้องมี)')

            # df_comment.loc[task_not_xx & task_new_na & comment_empty_or_na, 'COMMENT'] = 'ไม่มี TASK (ต้องมี)'
            # # df_comment.loc[task_not_xx & task_new_na & ~comment_empty_or_na, 'COMMENT'] += ', ไม่มี TASK (ต้องมี)'
            # df_comment.loc[task_not_xx & task_new_na & ~comment_empty_or_na, 'COMMENT'] = df_comment.loc[task_not_xx & task_new_na & ~comment_empty_or_na, 'COMMENT'].apply(
            #     lambda x: x.replace('ไม่มี TASK', 'ไม่มี TASK (ต้องมี)') if 'ไม่มี TASK' in x else f"{x}, ไม่มี TASK (ต้องมี)")
            replace_or_append_comment(df_comment, (task_not_xx & task_new_na), 'COMMENT', 'ไม่มี TASK (ต้องมี)', replace_message='ไม่มี TASK')

            # df_comment.loc[task_not_xx & start_date_na & comment_empty_or_na, 'COMMENT'] = 'ไม่มี START_DATE (ต้องมี)'
            # # df_comment.loc[task_not_xx & start_date_na & ~comment_empty_or_na, 'COMMENT'] += ', ไม่มี START_DATE (ต้องมี)'
            # df_comment.loc[task_not_xx & start_date_na & ~comment_empty_or_na, 'COMMENT'] = df_comment.loc[task_not_xx & start_date_na & ~comment_empty_or_na, 'COMMENT'].apply(
            #     lambda x: x.replace('ไม่มี START_DATE', 'ไม่มี START_DATE (ต้องมี)') if 'ไม่มี START_DATE' in x else f"{x}, ไม่มี START_DATE (ต้องมี)")
            replace_or_append_comment(df_comment, (task_not_xx & start_date_na), 'COMMENT', 'ไม่มี START_DATE (ต้องมี)', replace_message='ไม่มี START_DATE')


            # df_comment.loc[task_not_xx & finish_date_na & comment_empty_or_na, 'COMMENT'] = 'ไม่มี FINISH_DATE (ต้องมี)'
            # # df_comment.loc[task_not_xx & finish_date_na & ~comment_empty_or_na, 'COMMENT'] += ', ไม่มี FINISH_DATE (ต้องมี)'
            # df_comment.loc[task_not_xx & finish_date_na & ~comment_empty_or_na, 'COMMENT'] = df_comment.loc[task_not_xx & finish_date_na & ~comment_empty_or_na, 'COMMENT'].apply(
            #     lambda x: x.replace('ไม่มี FINISH_DATE', 'ไม่มี FINISH_DATE (ต้องมี)') if 'ไม่มี FINISH_DATE' in x else f"{x}, ไม่มี FINISH_DATE (ต้องมี)")
            replace_or_append_comment(df_comment, (task_not_xx & finish_date_na), 'COMMENT', 'ไม่มี FINISH_DATE (ต้องมี)', replace_message='ไม่มี FINISH_DATE')


            # df_comment.loc[task_not_xx & ptw_na & comment_empty_or_na, 'COMMENT'] = 'ไม่มี ประเภทของ_PERMIT_TO_WORK (ต้องมี)'
            # # df_comment.loc[task_not_xx & ptw_na & ~comment_empty_or_na, 'COMMENT'] += ', ไม่มี ประเภทของ_PERMIT_TO_WORK (ต้องมี)'
            # df_comment.loc[task_not_xx & ptw_na & ~comment_empty_or_na, 'COMMENT'] = df_comment.loc[task_not_xx & ptw_na & ~comment_empty_or_na, 'COMMENT'].apply(
            #     lambda x: x.replace('ไม่มี ประเภทของ_PERMIT_TO_WORK', 'ไม่มี ประเภทของ_PERMIT_TO_WORK (ต้องมี)') if 'ไม่มี ประเภทของ_PERMIT_TO_WORK' in x else f"{x}, ไม่มี ประเภทของ_PERMIT_TO_WORK (ต้องมี)")
            replace_or_append_comment(df_comment, (task_not_xx & ptw_na), 'COMMENT', 'ไม่มี ประเภทของ_PERMIT_TO_WORK (ต้องมี)', replace_message='ไม่มี ประเภทของ_PERMIT_TO_WORK')

            # df_comment.loc[task_not_xx & response_new_na & comment_empty_or_na, 'COMMENT'] = 'ไม่มี RESPONSE (ต้องมี)'
            # # df_comment.loc[task_not_xx & response_new_na & ~comment_empty_or_na, 'COMMENT'] += ', ไม่มี RESPONSE'
            # df_comment.loc[task_not_xx & response_new_na & ~comment_empty_or_na, 'COMMENT'] = df_comment.loc[task_not_xx & response_new_na & ~comment_empty_or_na, 'COMMENT'].apply(
            #     lambda x: x.replace('ไม่มี RESPONSE', 'ไม่มี RESPONSE (ต้องมี)') if 'ไม่มี RESPONSE' in x else f"{x}, ไม่มี RESPONSE (ต้องมี)")
            replace_or_append_comment(df_comment, (task_not_xx & response_new_na), 'COMMENT', 'ไม่มี RESPONSE (ต้องมี)', replace_message='ไม่มี RESPONSE')

            # df_comment.loc[task_not_xx & response_craft_na & comment_empty_or_na, 'COMMENT'] = 'ไม่มี RESPONSE_CRAFT (ต้องมี)'
            # # df_comment.loc[task_not_xx & response_craft_na & ~comment_empty_or_na, 'COMMENT'] += ', ไม่มี RESPONSE_CRAFT (ต้องมี)'
            # df_comment.loc[task_not_xx & response_craft_na & ~comment_empty_or_na, 'COMMENT'] = df_comment.loc[task_not_xx & response_craft_na & ~comment_empty_or_na, 'COMMENT'].apply(
            #     lambda x: x.replace('ไม่มี RESPONSE_CRAFT', 'ไม่มี RESPONSE_CRAFT (ต้องมี)') if 'ไม่มี RESPONSE_CRAFT' in x else f"{x}, ไม่มี RESPONSE_CRAFT (ต้องมี)")
            replace_or_append_comment(df_comment, (task_not_xx & response_craft_na), 'COMMENT', 'ไม่มี RESPONSE_CRAFT (ต้องมี)', replace_message='ไม่มี RESPONSE_CRAFT')

            # ลบคอมมาและช่องว่างที่ไม่จำเป็นออกจาก COMMENT (ถ้ามี)
            df_comment['COMMENT'] = df_comment['COMMENT'].str.strip(', ')


            #! Create Commnet.xlsx file
            df_comment['TASK_ORDER'] = df_original_copy['TASK_ORDER_NEW']
            red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
            yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
            blue_fill = PatternFill(start_color='00B0F0', end_color='00B0F0', fill_type='solid')
            center_alignment = Alignment(horizontal='center')

            df_comment = df_comment[['TASK_ORDER','COMMENT']]
            book = load_workbook(schedule_path)
            sheet = book.worksheets[0]
            # sheet = book[sheet_name]
            start_row = 3
            task_order_col = 7  # คอลัมน์เริ่มต้นสำหรับ TASK_ORDER
            comment_col = 19    # คอลัมน์เริ่มต้นสำหรับ COMMENT

            # เขียนข้อมูลลงใน Excel
            for i, (task_order, comment) in enumerate(zip(df_comment['TASK_ORDER'], df_comment['COMMENT']), start=start_row):
                # เขียนข้อมูล TASK_ORDER
                # task_order_cell = sheet.cell(row=i, column=task_order_col)
                # task_order_cell.value = task_order
                # task_order_cell.alignment = center_alignment  # จัดตัวอักษรให้อยู่ตรงกลาง
                # task_order_cell.fill = blue_fill
                
                # เขียนข้อมูล COMMENT
                comment_cell = sheet.cell(row=i, column=comment_col)
                comment_cell.value = comment
                
                if comment_cell.value:
                    if "(ต้องมี)" in comment_cell.value or "ไม่ถูกต้อง" in comment_cell.value:
                        comment_cell.fill = red_fill
                    else:
                        comment_cell.fill = yellow_fill

            book.save(comment_path)
            request.session['download_link_comment'] = comment_path
            
            #! RECHECK
            # ข้อมูลไม่สมบูรณ์
            task_order_not_xx = df_original['TASK_ORDER'] != 'xx'
            task_order = (df_original['TASK_ORDER'].notna()) & (df_original['TASK_ORDER'] != 'xx')
            task_not_xx = df_original_copy['TASK_ORDER'] != 'xx'
            # task_no_skill_rate = (df_original['TASK_ORDER'] == 10) & ((df_original['SUPERVISOR'].isna()) & (df_original['FOREMAN'].isna()) & (df_original['SKILL'].isna()))
            # task_no_type = (df_original['TASK_ORDER'].notna()) & (df_original['TASK_ORDER'] != 'xx') & (df_original['TYPE'].isna())
            # cond_duration = df_original['DURATION_(HR.)'].apply(lambda x: isinstance(x, (int, float)))
            # non_negative = df_original[cond_duration]['DURATION_(HR.)'] < 0
            # task_order_valid = df_original_check['TASK_ORDER_NEW'].apply(
            #     lambda x: (isinstance(x, str) and (x.isdigit() or x == 'xx')) or 
            #             (isinstance(x, (int, float)) and x.is_integer() and x >= 0)
            # )            
            # cond_start_date_new = df_original['START_DATE'].notna() & df_original['START_DATE_NEW'].isna() & (df_original['TASK_ORDER'].notna()) & (df_original['TASK_ORDER'] != 'xx')
            # cond_finish_date_new = df_original['FINISH_DATE'].notna() & df_original['FINISH_DATE_NEW'].isna() & (df_original['TASK_ORDER'].notna()) & (df_original['TASK_ORDER'] != 'xx')
            # valid_supervisor = df_original['SUPERVISOR'].apply(lambda x: (pd.isna(x) or (isinstance(x, (int, float)) and x.is_integer() and x >= 0)))
            # valid_foreman = df_original['FOREMAN'].apply(lambda x: (pd.isna(x) or (isinstance(x, (int, float)) and x.is_integer() and x >= 0)))
            # valid_skill = df_original['SKILL'].apply(lambda x: (pd.isna(x) or (isinstance(x, (int, float)) and x.is_integer() and x >= 0)))
            # # เงื่อนไขตรวจสอบว่าค่าใน SUPERVISOR, FOREMAN, SKILL มีค่าที่ไม่ถูกต้อง (ไม่เป็นจำนวนเต็มบวก)
            # invalid_skill_rate = ~(valid_supervisor & valid_foreman & valid_skill)
            # task_type = (df_original['TASK_ORDER'].notna()) & (df_original['TASK_ORDER'] != 'xx') & (df_original['TYPE'].notna())
            # valid_type = (df_original['TYPE'].isin(['ME', 'EE', 'CV', 'IC']))
            
            kks_new_na = df_original_copy['KKS_NEW'].isin([''])
            equip_new_na = df_original_copy['EQUIPMENT_NEW'].isin([''])
            task_order_new_na = df_original_copy['TASK_ORDER_NEW'].isin([''])
            task_new_na = df_original_copy['TASK_NEW'].isin([''])
            start_date_na = df_original_copy['START_DATE'].isin([''])
            finish_date_na = df_original_copy['FINISH_DATE'].isin([''])
            ptw_na = df_original_copy['ประเภทของ_PERMIT_TO_WORK'].isin([''])
            response_new_na = df_original_copy['RESPONSE'].isin([''])
            response_craft_na = df_original_copy['RESPONSE_CRAFT'].isin([''])

            missing_columns = []
            # ตรวจสอบเงื่อนไขแต่ละคอลัมน์และเก็บชื่อคอลัมน์ที่ข้อมูลขาดหาย
            if ((task_not_xx) & (kks_new_na)).any():
                missing_columns.append('KKS')
            if ((task_not_xx) & (equip_new_na)).any():
                missing_columns.append('EQUIPMENT')
            if ((task_not_xx) & (task_order_new_na)).any():
                missing_columns.append('TASK_ORDER')
            if ((task_not_xx) & (task_new_na)).any():
                missing_columns.append('TASK')
            if ((task_not_xx) & (start_date_na)).any():
                missing_columns.append('START_DATE')
            if ((task_not_xx) & (finish_date_na)).any():
                missing_columns.append('FINISH_DATE')
            if ((task_not_xx) & (ptw_na)).any():
                missing_columns.append('PTW')
            if ((task_not_xx) & (response_new_na)).any():
                missing_columns.append('RESPONSE')
            if ((task_not_xx) & (response_craft_na)).any():
                missing_columns.append('RESPONSE_CRAFT')
            if (task_no_skill_rate).any():
                missing_columns.append('SKILL RATE')
            if (task_no_type).any():
                missing_columns.append('TYPE')

            invalid_columns = []
            # ตรวจสอบเงื่อนไขแต่ละคอลัมน์และเก็บชื่อคอลัมน์ที่ข้อมูลไม่ถูกต้อง
            if ((~cond_duration | non_negative) & task_order_not_xx).any():
                invalid_columns.append('DURATION_(HR.)')
            if (~task_order_valid & (df_original_check['TASK_ORDER']!= -1)).any():
                invalid_columns.append('TASK_ORDER')
            if (cond_start_date_new).any():
                invalid_columns.append('START_DATE')
            if (cond_finish_date_new).any():
                invalid_columns.append('FINISH_DATE')
            if ((task_order) & (invalid_skill_rate)).any():
                invalid_columns.append('SKILL RATE')
            if ((task_type) & (~valid_type)).any():
                invalid_columns.append('TYPE')
            
            # การตรวจสอบข้อมูล
            if missing_columns or invalid_columns:
                if missing_columns:
                    error_message += f"พบข้อมูลขาดหายในคอลัมน์: {', '.join(missing_columns)} "
                if invalid_columns:
                    error_message += f"พบข้อมูลที่ไม่ถูกต้องในคอลัมน์: {', '.join(invalid_columns)}"
                
                # ส่งข้อความแจ้งเตือนไปยังหน้า upload.html
                return render(request, 'maximo_app/upload.html', {
                    'error_message': error_message,
                    'schedule_filename': schedule_filename,
                    'location_filename': location_filename,
                })
            #! END RECHECK
############
############
            #? extracted_kks_counts
            df_original_newcol['UNIT'] = df_original_newcol['KKS_NEW'].str[0:3]
            df_original_extracted['EXTRACTED_KKS'] = df_original_extracted['KKS'].str.extract(r'-(\w{3})')
            extracted_kks_counts = df_original_extracted['EXTRACTED_KKS'].value_counts()
############
############
            timestamp_columns_1 = ['START_DATE', 'FINISH_DATE']
            timestamp_columns_2 = ['START_DATE_NEW', 'FINISH_DATE_NEW']
            
            # Timestamp to String
            df_original = convert_timestamp_columns_to_str(df_original, timestamp_columns_2)
            df_original_copy = convert_timestamp_columns_to_str(df_original_copy, timestamp_columns_1)
            df_original_copy = convert_timestamp_columns_to_str(df_original_copy, timestamp_columns_2)
            df_original_newcol = convert_timestamp_columns_to_str(df_original_newcol, timestamp_columns_1)
            
            # Save DataFrames
            request.session['df_original'] = df_original.to_dict(orient='index')
            request.session['df_original_copy'] = df_original_copy.to_dict(orient='index')
            request.session['df_original_newcol'] = df_original_newcol.to_dict(orient='index')
            request.session['df_comment'] = df_comment.to_dict(orient='index')
            request.session['extracted_kks_counts'] = extracted_kks_counts.to_dict()
            
            # Save Series
            # task_no_skill_rate: DataFrame, ใช้ orient='index'
            # request.session['task_no_skill_rate'] = task_no_skill_rate.to_dict()
            # request.session['task_no_type'] = task_no_type.to_dict()
            # request.session['cond_duration'] = cond_duration.to_dict()
            # request.session['non_negative'] = non_negative.to_dict()
            # request.session['task_order_valid'] = task_order_valid.to_dict()
            # request.session['cond_start_date_new'] = cond_start_date_new.to_dict()
            # request.session['cond_finish_date_new'] = cond_finish_date_new.to_dict()
            # request.session['invalid_skill_rate'] = invalid_skill_rate.to_dict()
            # request.session['task_type'] = task_type.to_dict()
            # request.session['valid_type'] = valid_type.to_dict()
            
            # Save variables
            request.session['first_plant'] = first_plant
            request.session['most_common_plant_unit'] = most_common_plant_unit
############
############        
            print('1 form schedule_filename:',schedule_filename)
            print('1 form location_filename:',location_filename)
            print('1 form comment_path:',comment_path)
            print('1 form extracted_kks_counts:',extracted_kks_counts)
            print('1 form user_input_mapping:',user_input_mapping)
############
############                        
        elif 'kks_mapping_submit' in request.POST:
            # Get variables
            schedule_filename = request.session.get('schedule_filename', '')
            location_filename = request.session.get('location_filename', '')
            comment_path = request.session.get('download_link_comment', None)
            first_plant = request.session.get('first_plant')
            temp_dir = request.session.get('temp_dir')
            most_common_plant_unit = request.session.get('most_common_plant_unit')
            # Get Dropdown
            egmntacttype = request.session.get('egmntacttype')
            egprojectid = request.session.get('egprojectid')
            egwbs = request.session.get('egwbs')
            location = request.session.get('location')
            siteid = request.session.get('siteid')
            wbs_desc = request.session.get('wbs_desc')
            worktype = request.session.get('worktype')
            wostatus = request.session.get('wostatus')
            
            # Get DataFrames
            extracted_kks_counts = pd.DataFrame.from_dict(request.session.get('extracted_kks_counts', {}), orient='index')
            df_original = pd.DataFrame.from_dict(request.session['df_original'], orient='index')
            df_original_copy = pd.DataFrame.from_dict(request.session['df_original_copy'], orient='index')
            df_original_newcol = pd.DataFrame.from_dict(request.session['df_original_newcol'], orient='index')
            df_comment = pd.DataFrame.from_dict(request.session['df_comment'], orient='index')
            
            # Get Series
            # task_no_skill_rate = pd.DataFrame.from_dict(request.session['task_no_skill_rate'], orient='index')
            # task_no_type = pd.DataFrame.from_dict(request.session['task_no_type'], orient='index')
            # cond_duration = pd.DataFrame.from_dict(request.session['cond_duration'], orient='index')
            # non_negative = pd.DataFrame.from_dict(request.session['non_negative'], orient='index')
            # task_order_valid = pd.DataFrame.from_dict(request.session['task_order_valid'], orient='index')
            # cond_start_date_new = pd.DataFrame.from_dict(request.session['cond_start_date_new'], orient='index')
            # cond_finish_date_new = pd.DataFrame.from_dict(request.session['cond_finish_date_new'], orient='index')
            # invalid_skill_rate = pd.DataFrame.from_dict(request.session['invalid_skill_rate'], orient='index')
            # task_type = pd.DataFrame.from_dict(request.session['task_type'], orient='index')
            # valid_type = pd.DataFrame.from_dict(request.session['valid_type'], orient='index')
            
            # Define path
            job_plan_task_path = os.path.join(temp_dir, f"{uuid.uuid4()}_Job_Plan.xlsx")
            job_plan_labor_path = os.path.join(temp_dir, f"{uuid.uuid4()}_Job_Plan_Labor.xlsx")
            pm_plan_path = os.path.join(temp_dir, f"{uuid.uuid4()}_PM_Plan.xlsx")
            

            timestamp_columns_1 = ['START_DATE', 'FINISH_DATE']
            timestamp_columns_2 = ['START_DATE_NEW', 'FINISH_DATE_NEW']
            df_original = convert_str_columns_to_timestamp(df_original, timestamp_columns_2)
            df_original_copy = convert_str_columns_to_timestamp(df_original_copy, timestamp_columns_1)
            df_original_copy = convert_str_columns_to_timestamp(df_original_copy, timestamp_columns_2)
            df_original_newcol = convert_str_columns_to_timestamp(df_original_newcol, timestamp_columns_1)
            
            # Collect mapping data from the POST request
            if not extracted_kks_counts.empty:
                user_input_mapping = {}
                for kks_value in extracted_kks_counts.index:
                    user_value = request.POST.get(kks_value)  # รับค่าจาก text box ในฟอร์ม
                    user_input_mapping[kks_value] = user_value  # เก็บใน dictionary
                request.session['user_input_mapping'] = user_input_mapping
                # นำข้อมูล user_input_mapping ไปใช้ในขั้นตอนถัดไป
            else:
                return HttpResponse("Error: KKS counts data is missing.")
############
############        
            print('2 elif schedule_filename:',schedule_filename)
            print('2 elif location_filename:',location_filename)
            print('2 elif comment_path:',  comment_path)
            print('2 elif extracted_kks_counts:',extracted_kks_counts)
            print('2 elif user_input_mapping:',user_input_mapping)
############
############  
            orgid = 'EGAT'
            pluscrevum = 0
            status = 'ACTIVE'
            pluscjprevnum = 0
            frequency = 4
            frequnit = 'YEARS'
            route = ''
            leadtime = 7
############
############
            #! Creat JOB PLAN TASK

            # user_input_mapping = {}
            # for kks_value in extracted_kks_counts.index:
            #     user_value = input((f"Please enter the mapping for '{kks_value}': "))
            #     user_input_mapping[kks_value] = user_value

            # df_original_newcol['UNIT_TYPE']= df_original_newcol['UNIT'].map({'H00':'C0', 'H13':'C13', 'H03':'U', 'H04':'U'})
            df_original_newcol['UNIT_TYPE'] = df_original_newcol['UNIT'].map(user_input_mapping)


            df_original_filter = df_original_newcol[df_original_newcol['TASK_ORDER_NEW']!='xx'].copy()
            # for i,j in zip(df_original_filter['TASK_ORDER_NEW'], df_original_filter.index):
            #     if (type(i)!=float) and (type(i)!=int):
            #         i,j
            df_original_filter['TASK_ORDER_NEW'] = df_original_filter['TASK_ORDER_NEW'].astype('int32')

            cond1 = (df_original_filter['TASK_ORDER_NEW']==10) & (df_original_filter['TASK_ORDER_NEW'].shift(1)!=10)
            cond2 = ((df_original_filter['TASK_ORDER_NEW']==10) & 
            (df_original_filter['TASK_ORDER_NEW'].shift(1)==10) & (df_original_filter['DURATION_(HR.)']!=-1))
            xx = cond1|cond2

            df_original_filter['GROUP_LEVEL_1'] = xx.cumsum().rename('GROUP_LEVEL_1')
            # df_original_filter['GROUP_LEVEL_2'] = df_original_filter['GROUP_LEVEL_1'].astype('str')+'-'+df_original_filter['KKS_NEW']
            df_original_filter['GROUP_LEVEL_2'] = (df_original_filter['GROUP_LEVEL_1'].astype(int) * 10).astype(str).str.zfill(5) + '-' + df_original_filter['KKS_NEW']
            df_original_filter['GROUP_LEVEL_2'] = df_original_filter['GROUP_LEVEL_2'].astype(str)
            df_original_filter['TYPE'] = df_original_filter['TYPE'].astype(str)
            df_original_filter['UNIT_TYPE'] = df_original_filter['UNIT_TYPE'].astype(str)
            df_original_filter['GROUP_LEVEL_3'] = df_original_filter['GROUP_LEVEL_2']+'-'+df_original_filter['TYPE']+'-'+df_original_filter['UNIT_TYPE']

            common_indices = df_original_copy.index.intersection(df_original_filter.index)
            df_original_copy.loc[common_indices, ['GROUP_LEVEL_1', 'GROUP_LEVEL_2','GROUP_LEVEL_3']] = df_original_filter.loc[common_indices, ['GROUP_LEVEL_1', 'GROUP_LEVEL_2','GROUP_LEVEL_3']]
            common_indices1 = df_original.index.intersection(df_original_filter.index)
            df_original.loc[common_indices1, ['GROUP_LEVEL_1', 'GROUP_LEVEL_2','GROUP_LEVEL_3']] = df_original_filter.loc[common_indices1, ['GROUP_LEVEL_1', 'GROUP_LEVEL_2','GROUP_LEVEL_3']]

            data = df_original_filter.groupby(['GROUP_LEVEL_1','GROUP_LEVEL_3',
                                            'KKS_NEW','EQUIPMENT_NEW']).size()
            df_original_filter_group = pd.DataFrame(data).reset_index()
            df_original_filter_group = df_original_filter_group.drop(columns = [0])

            dict_jp_master = {}
            lst = ['GROUP_LEVEL_1','GROUP_LEVEL_3','KKS_NEW','EQUIPMENT_NEW','DURATION_TOTAL','TASK_ORDER_NEW', 'DURATION_(HR.)',
                'START_DATE','FINISH_DATE','TASK_NEW','RESPONSE_CRAFT', 'ประเภทของ_PERMIT_TO_WORK']
            #####################
            # for group_1 in [4,10,22]: # for testing
            for group_1 in df_original_filter_group['GROUP_LEVEL_1'].value_counts().sort_index().index:
                cond1 = (df_original_filter['GROUP_LEVEL_1']==group_1)
                ############
                work_group = df_original_filter[cond1].iloc[:,0:].reset_index(drop=True) # separate each group
                ###########
                cond2 = work_group['DURATION_(HR.)']!=-1 # to eliminate redundant task
                df_new = work_group[cond2].copy()
                df_new['DURATION_TOTAL'] = df_new['DURATION_(HR.)'].sum()
                df_new = df_new.replace(-1,np.nan)
                df_new = df_new.ffill()
                df_new = df_new.fillna(-1)
                dict_jp_master[group_1] = df_new[lst]

            dict_jp_master2 = {}
            df_jop_plan = pd.DataFrame()
            for group1,group3,eq in zip(df_original_filter_group['GROUP_LEVEL_1'],
                                        df_original_filter_group['GROUP_LEVEL_3'],
                                        df_original_filter_group['EQUIPMENT_NEW']):
            #     group1,group2,eq
                df = dict_jp_master[group1].copy()
                lst = ['DURATION_TOTAL','TASK_ORDER_NEW','DURATION_(HR.)','START_DATE','FINISH_DATE','TASK_NEW']
                df_new = df[lst].copy()
                # ขยายค่า group3 และ eq ให้เท่ากับจำนวนแถวของ df_new
                df_new.loc[:, 'JOB_NUM'] = [group3] * len(df_new)
                df_new.loc[:, 'EQUIPMENT'] = [eq] * len(df_new)
                lst2 = ['JOB_NUM','EQUIPMENT','DURATION_TOTAL','TASK_ORDER_NEW','DURATION_(HR.)','START_DATE','FINISH_DATE','TASK_NEW']
                dict_jp_master2[group3]=df_new[lst2]
                df_jop_plan = pd.concat([df_jop_plan,df_new[lst2]])
            df_jop_plan = df_jop_plan.reset_index(drop=True)
            df_jop_plan['GROUP_NUM'] = df_jop_plan.groupby('JOB_NUM').ngroup()+1
            df1 = df_jop_plan['GROUP_NUM'].drop_duplicates()
            df1 = df1.reset_index(drop=True)
            df2 = pd.Series(np.arange(1,len(df1)+1),name = 'MAP_GROUP')
            df3 = pd.concat([df2,df1],axis=1)
            df3 = df3.set_index('GROUP_NUM')
            dict_new = df3.to_dict()['MAP_GROUP']
            df_jop_plan['GROUP']=df_jop_plan['GROUP_NUM'].map(dict_new)

            df_jop_plan['ORGID']= orgid
            df_jop_plan['SITEID']= siteid
            df_jop_plan['PLUSCREVNUM']= pluscrevum
            df_jop_plan['STATUS']= status
            df_jop_plan['PLUSCJPREVNUM']= pluscjprevnum

            new_columns= ['JOB_NUM', 'ORGID','SITEID','PLUSCREVNUM','STATUS',
                                'EQUIPMENT', 'DURATION_TOTAL', 'TASK_ORDER_NEW','PLUSCJPREVNUM',
                                'DURATION_(HR.)', 'TASK_NEW','START_DATE', 'FINISH_DATE',
                                'GROUP']
            df_jop_plan_master = df_jop_plan[new_columns].copy()
            df_jop_plan_master['MOD'] = df_jop_plan_master['GROUP']%2
            df_jop_plan_master['JOB_NUM'] = 'JP'+'-'+df_jop_plan_master['JOB_NUM']

            df_jop_plan_master['COMMENT'] = ''
            job_plan_cond1 = df_jop_plan_master['JOB_NUM'].str.len()>30
            # comment_empty_or_na = (df_jop_plan_master['COMMENT'] == '') | (df_jop_plan_master['COMMENT'].isna())
            # df_jop_plan_master.loc[job_plan_cond1 & comment_empty_or_na, 'COMMENT'] = 'JPNUM มีความยาวมากกว่า 30 ตัวอักษร'
            # df_jop_plan_master.loc[job_plan_cond1 & ~comment_empty_or_na, 'COMMENT'] += ', JPNUM มีความยาวมากกว่า 30 ตัวอักษร'
            update_comment(df_jop_plan_master, job_plan_cond1, 'COMMENT', 'JPNUM มีความยาวมากกว่า 30 ตัวอักษร')


            job_plan_cond2 = df_jop_plan_master['EQUIPMENT'].str.len()>100
            # comment_empty_or_na = (df_jop_plan_master['COMMENT'] == '') | (df_jop_plan_master['COMMENT'].isna())
            # df_jop_plan_master.loc[job_plan_cond2 & comment_empty_or_na, 'COMMENT'] = 'DESCRIPTION มีความยาวมากกว่า 100 ตัวอักษร'
            # df_jop_plan_master.loc[job_plan_cond2 & ~comment_empty_or_na, 'COMMENT'] += ', DESCRIPTION มีความยาวมากกว่า 100 ตัวอักษร'
            update_comment(df_jop_plan_master, job_plan_cond2, 'COMMENT', 'DESCRIPTION มีความยาวมากกว่า 100 ตัวอักษร')

            job_plan_cond3 = df_jop_plan_master['TASK_NEW'].str.len()>100
            # comment_empty_or_na = (df_jop_plan_master['COMMENT'] == '') | (df_jop_plan_master['COMMENT'].isna())
            # df_jop_plan_master.loc[job_plan_cond3 & comment_empty_or_na, 'COMMENT'] = 'JOBTASK มีความยาวมากกว่า 100 ตัวอักษร'
            # df_jop_plan_master.loc[job_plan_cond3 & ~comment_empty_or_na, 'COMMENT'] += ', JOBTASK มีความยาวมากกว่า 100 ตัวอักษร'
            update_comment(df_jop_plan_master, job_plan_cond3, 'COMMENT', 'JOBTASK มีความยาวมากกว่า 100 ตัวอักษร')

            #! Create Job_Plan_Task.xlsx
            df_jop_plan_master.to_excel(job_plan_task_path,index=False)


            #! JOB PLAN Labor
            df_jop_plan_master_labor = df_jop_plan_master.copy()
            df_jop_plan_master_labor['GROUP_LEVEL_1'] = df_jop_plan_master_labor['JOB_NUM'].str.extract(r'JP-(\d+)-')
            df_jop_plan_master_labor['GROUP_LEVEL_1'] = df_jop_plan_master_labor['GROUP_LEVEL_1'].astype('int32')
            df_jop_plan_master_labor[['JOB_NUM','DURATION_TOTAL','GROUP_LEVEL_1']]

            df_group_level_1 = pd.DataFrame()
            for group_1 in df_original_filter_group['GROUP_LEVEL_1'].value_counts().sort_index().index:## get group 1
                cond1 = (df_original_filter['GROUP_LEVEL_1']==group_1)
                ############
                work_group = df_original_filter[cond1].iloc[:,0:].reset_index(drop=True) # separate each group
                ###########
                cond2 = work_group['DURATION_(HR.)']!=-1 # to eliminate redundant task
                df_new = work_group[cond2].copy()
                df_new['DURATION_TOTAL']=df_new['DURATION_(HR.)'].sum()
                df_new = df_new.replace(-1,np.nan)
                df_new = df_new.ffill()
                df_new = df_new.fillna(-1)
                lst = ['RESPONSE_CRAFT','SUPERVISOR','FOREMAN','SKILL','DURATION_TOTAL','GROUP_LEVEL_1']
                df_group_level_1 = pd.concat([df_group_level_1,df_new[lst]])

            df_group_level_1 = df_group_level_1.reset_index(drop=True)

            df_labor = pd.DataFrame()
            for group_level_1 in df_jop_plan_master_labor['GROUP_LEVEL_1'].value_counts().sort_index().index:
                df1 = df_jop_plan_master_labor[df_jop_plan_master_labor['GROUP_LEVEL_1']==group_level_1][['JOB_NUM','DURATION_TOTAL','GROUP_LEVEL_1']]
                df2 = df_group_level_1[df_group_level_1['GROUP_LEVEL_1']==group_level_1]
                df1 = df1.reset_index(drop=True)
                df2 = df2.reset_index(drop=True)
                dfx = pd.concat([df1,df2],axis=1)
                _, i = np.unique(dfx.columns, return_index=True) ## Eliminate redundant columns
                i.sort()
                dfxx = dfx.iloc[:, i].copy()
                dfxx = dfxx.rename(columns={"SUPERVISOR": "L21NOM", "FOREMAN": "L22NOM", "SKILL":"L23NOM"})
                #dfxx.iloc[0:,0:]
                df_labor = pd.concat([df_labor,dfxx.iloc[0:1,0:]])

            df_labor = df_labor.reset_index(drop=True)
            df_labor.replace(-1,np.nan, inplace=True)
            df_melt = pd.melt(df_labor, id_vars=['JOB_NUM','DURATION_TOTAL','GROUP_LEVEL_1','RESPONSE_CRAFT'], value_vars=['L21NOM', 'L22NOM','L23NOM'])
            df_labor = df_melt.sort_values(by = ['GROUP_LEVEL_1','variable']).dropna()
            df_labor = df_labor.reset_index(drop=True)
            lst = ['JOB_NUM', 'RESPONSE_CRAFT', 'variable','DURATION_TOTAL', 'value']
            df_labor_new = df_labor[lst].copy()
            df_labor_new.columns = ['JOB_NUM', 'CRAFT', 'SKILLLEVEL', 'LABORHRS', 'QUANTITY']
            df_labor_new['GROUP_NUM'] = df_labor_new.groupby('JOB_NUM').ngroup()+1
            df1 = df_labor_new['GROUP_NUM'].drop_duplicates()
            df1 = df1.reset_index(drop=True)
            df2 = pd.Series(np.arange(1,len(df1)+1),name = 'MAP_GROUP')
            df3 = pd.concat([df2,df1],axis=1)
            df3 = df3.set_index('GROUP_NUM')
            dict_new = df3.to_dict()['MAP_GROUP']
            df_labor_new['GROUP']=df_labor_new['GROUP_NUM'].map(dict_new)
            df_labor_new['MOD'] = df_labor_new['GROUP']%2

            df_labor_new['ORGID'] = orgid
            df_labor_new['SITEID'] = siteid
            df_labor_new['PLUSCREVNUM']= pluscrevum
            df_labor_new['STATUS']= status
            df_labor_new['JPTASK']= ''

            df_labor_new.columns
            lst_labor1 = ['JOB_NUM','ORGID','SITEID','PLUSCREVNUM', 
                'STATUS','JPTASK', 'CRAFT','SKILLLEVEL',
                'LABORHRS','QUANTITY','GROUP','MOD', 'COMMENT']

            df_labor_new['COMMENT'] = ''
            job_plan_cond1 = df_labor_new['JOB_NUM'].str.len()>30
            # comment_empty_or_na = (df_labor_new['COMMENT'] == '') | (df_labor_new['COMMENT'].isna())
            # df_labor_new.loc[job_plan_cond1 & comment_empty_or_na, 'COMMENT'] = 'JPNUM มีความยาวมากกว่า 30 ตัวอักษร'
            # df_labor_new.loc[job_plan_cond1 & ~comment_empty_or_na, 'COMMENT'] += ', JPNUM มีความยาวมากกว่า 30 ตัวอักษร'
            update_comment(df_labor_new, job_plan_cond1, 'COMMENT', 'JPNUM มีความยาวมากกว่า 30 ตัวอักษร')

            #! Create Job_Plan_Labor.xlsx
            df_labor_new[lst_labor1].to_excel(job_plan_labor_path, index=False)


            #! Create PM PLAN
            data_pm = df_original_filter.groupby(['GROUP_LEVEL_1','GROUP_LEVEL_3','KKS_NEW','MAIN_SYSTEM','SUB_SYSTEM',
                                        'EQUIPMENT','MAIN_SYSTEM_DESC','SUB_SYSTEM_DESC','KKS_NEW_DESC','EQUIPMENT_NEW',
                                        'UNIT_TYPE','TYPE'
                                    ]).size()
            df_original_filter_group_pm = pd.DataFrame(data_pm).reset_index()
            df_original_filter_group_pm = df_original_filter_group_pm.drop(columns = [0])

            for i,df in df_original_filter_group_pm.iloc[20:22].iterrows():
                i,df

            pm_master_dict = {'PMNUM':[],
                            'SITEID':[],
                            'DESCRIPTION':[],
                            'STATUS':[],
                            'LOCATION':[],
                            'ROUTE':[],
                            'LEADTIME':[],
                            'PMCOUNTER':[],
                            'WORKTYPE':[],
                            'EGMNTACTTYPE':[],
                            'WOSTATUS':[],
                            'EGCRAFT':[],
                            'RESPONSED BY':[],
                            'PTW':[],
                            'LOTO':[],
                            'EGPROJECTID':[],
                            'EGWBS':[],
                            'FREQUENCY':[],
                            'FREQUNIT':[],
                            'NEXTDATE':[],
                            'TARGSTARTTIME':[],
                            'FINISH_DATE':[],
                            'FINISH TIME':[],
                            'PARENT':[],
                            'JPNUM':[],
                            'MAIN_SYSTEM':[],
                            'SUB_SYSTEM':[],
                            'EQUIPMENT':[],
                            'MAIN_SYSTEM_DESC':[],
                            'SUB_SYSTEM_DESC':[],
                            'KKS_NEW_DESC':[],
                            'UNIT_TYPE':[],
                            'TYPE':[],
                            }
            ############################
            for row_index,df in  df_original_filter_group_pm.iterrows():
                #row_index
                ####PM Number
                df_group_temp = df_original_filter.copy()
                PMNUM = 'PO'+'-'+df['GROUP_LEVEL_3']
                pm_master_dict['PMNUM'].append(PMNUM)
                #####Site ID
                pm_master_dict['SITEID'] = siteid
                #####DESCRIPTION
                desc = df['EQUIPMENT_NEW']
                pm_master_dict['DESCRIPTION'].append(desc)
                #####STATUS
                pm_master_dict['STATUS'] = status
                #####LOCATION
                loc = first_plant+df['KKS_NEW']
                pm_master_dict['LOCATION'].append(loc)
                #####ROUTE
                pm_master_dict['ROUTE'] = route
                #####LEADTIME
                pm_master_dict['LEADTIME']= leadtime
                #####PMCOUNTER
                pm_master_dict['PMCOUNTER']= ''
                #####WORKTYPE
                pm_master_dict['WORKTYPE']= worktype
                #####EGMNTACTTYPE
                pm_master_dict['EGMNTACTTYPE']= egmntacttype
                #####WOSTATUS
                pm_master_dict['WOSTATUS'] = wostatus
                ######EGCRAFT
                carft_lst = df_group_temp[df_group_temp['GROUP_LEVEL_1']==df['GROUP_LEVEL_1']]['RESPONSE_CRAFT'].drop_duplicates().to_list()
                carft = '//'.join(carft_lst)
                pm_master_dict['EGCRAFT'].append(carft)
                ######RESPONSED BY
                response_lst = df_group_temp[df_group_temp['GROUP_LEVEL_1']==df['GROUP_LEVEL_1']]['RESPONSE'].drop_duplicates().to_list()
                response = '//'.join(response_lst)
                pm_master_dict['RESPONSED BY'].append(response)
                ######PTW
                ptw_lst = df_group_temp[df_group_temp['GROUP_LEVEL_1']==df['GROUP_LEVEL_1']]['ประเภทของ_PERMIT_TO_WORK'].drop_duplicates().to_list()
                ptw = '//'.join(ptw_lst)
                pm_master_dict['PTW'].append(ptw)
                ######LOTO
                pm_master_dict['LOTO']= ''
                ######EGPROJECTID
                pm_master_dict['EGPROJECTID'] = egprojectid
                ######EGWBS
                pm_master_dict['EGWBS']= egwbs
                ######FREQUENCY
                pm_master_dict['FREQUENCY']= frequency
                ######FREQUNIT
                pm_master_dict['FREQUNIT']= frequnit
                ######NEXTDATE
                next_date = df_group_temp[df_group_temp['GROUP_LEVEL_1']==df['GROUP_LEVEL_1']]['START_DATE'].min()
                pm_master_dict['NEXTDATE'].append(next_date.date())
                ######TARGSTARTTIME
                time_start = datetime.time(8,0)
                pm_master_dict['TARGSTARTTIME'] = time_start
                #######FINISH_DATE
                last_date = df_group_temp[df_group_temp['GROUP_LEVEL_1']==df['GROUP_LEVEL_1']]['FINISH_DATE'].max()
                pm_master_dict['FINISH_DATE'].append(last_date.date())
                #######FINISH TIME
                time_end = datetime.time(16,0)
                pm_master_dict['FINISH TIME'] = time_end
                #######PARENT
                pm_master_dict['PARENT']=''
                #######JPNUM
                JPNUM = 'JP'+'-'+df['GROUP_LEVEL_3']
                pm_master_dict['JPNUM'].append(JPNUM)
                #######main_system
                main_system = df['MAIN_SYSTEM']
                pm_master_dict['MAIN_SYSTEM'].append(main_system)
                #######sub_system
                sub_system = df['SUB_SYSTEM']
                pm_master_dict['SUB_SYSTEM'].append(sub_system)
                #######equipment
                equipment = df['EQUIPMENT']
                pm_master_dict['EQUIPMENT'].append(equipment)
                #######main_system_desc
                main_system_desc = df['MAIN_SYSTEM_DESC']
                pm_master_dict['MAIN_SYSTEM_DESC'].append(main_system_desc)
                #######sub_system_desc
                sub_system_desc = df['SUB_SYSTEM_DESC']
                pm_master_dict['SUB_SYSTEM_DESC'].append(sub_system_desc)
                #######kks_new_desc
                kks_new_desc = df['KKS_NEW_DESC']
                pm_master_dict['KKS_NEW_DESC'].append(kks_new_desc)
                #######unit_type
                unit_type = df['UNIT_TYPE']
                pm_master_dict['UNIT_TYPE'].append(unit_type)
                #######carft_desc
                TYPE = df['TYPE']
                pm_master_dict['TYPE'].append(TYPE)

            pm_master_df = pd.DataFrame(pm_master_dict)
            
            # pm_master_df1['PTW'] = pm_master_df['PTW'].apply(clean_ptw_column)
            pm_master_df['PTW'] = pm_master_df['PTW'].apply(clean_ptw_column)
            #? END Clean
            
            pm_master_df[['SUB_SYSTEM','SUB_SYSTEM_DESC','KKS_NEW_DESC','EQUIPMENT','UNIT_TYPE','TYPE']].drop_duplicates().sort_values(by=['TYPE','UNIT_TYPE'],ascending=False)

            df_yy = pm_master_df.groupby(['MAIN_SYSTEM','MAIN_SYSTEM_DESC','EGCRAFT','PTW','UNIT_TYPE','TYPE']).size().sort_index(level=[5,4,3],ascending=False)
            df_yyy = df_yy.reset_index()
            df_yyy.rename(columns={0: "count"}, inplace=True)

            df_yyy_cont_more = df_yyy[df_yyy['count']>1].reset_index(drop=True).copy()
            df_yyy_cont_less = df_yyy[df_yyy['count']<=1].reset_index(drop=True).copy()

            df_pm_plan3_more = pd.DataFrame()
            loop_num = 0
            for index,row in df_yyy_cont_more.iterrows():
                loop_num +=1 
                #################
                group = pm_master_df[(pm_master_df['MAIN_SYSTEM']==row['MAIN_SYSTEM']) &
                                    (pm_master_df['MAIN_SYSTEM_DESC']==row['MAIN_SYSTEM_DESC']) &
                                    (pm_master_df['EGCRAFT']==row['EGCRAFT'])&
                                    (pm_master_df['PTW']==row['PTW'])&
                                    (pm_master_df['UNIT_TYPE']==row['UNIT_TYPE']) &
                                    (pm_master_df['TYPE']==row['TYPE'])]
                #group
                                    
                ###################
                ###################
                pm_parent_dict = {
                            'PMNUM':[],
                            'SITEID':[],
                            'DESCRIPTION':[],
                            'STATUS':[],
                            'LOCATION':[],
                            'ROUTE':[],
                            'LEADTIME':[],
                            'PMCOUNTER':[],
                            'WORKTYPE':[],
                            'EGMNTACTTYPE':[],
                            'WOSTATUS':[],
                            'EGCRAFT':[],
                            'RESPONSED BY':[],
                            'PTW':[],
                            'LOTO':[],
                            'EGPROJECTID':[],
                            'EGWBS':[],
                            'FREQUENCY':[],
                            'FREQUNIT':[],
                            'NEXTDATE':[],
                            'TARGSTARTTIME':[],
                            'FINISH_DATE':[],
                            'FINISH TIME':[],
                            'PARENT':[],
                            'JPNUM':[],
                            'MAIN_SYSTEM':[],
            #               'SUB_SYSTEM':[],
            #               'EQUIPMENT':[],
                            'MAIN_SYSTEM_DESC':[],
            #               'SUB_SYSTEM_DESC':[],
                            'UNIT_TYPE':[],
                            'TYPE':[]
                            }
                if loop_num < 10:
                    pm_parent_dict['PMNUM'] = 'Group-0{}-{}-{}-{}'.format(loop_num,row['MAIN_SYSTEM'],row['TYPE'],row['UNIT_TYPE'])
                else:
                    pm_parent_dict['PMNUM'] = 'Group-{}-{}-{}-{}'.format(loop_num,row['MAIN_SYSTEM'],row['TYPE'],row['UNIT_TYPE'])
                    
                pm_parent_dict['SITEID'] = siteid
                pm_parent_dict['DESCRIPTION'] = '{},{},{}'.format(row['MAIN_SYSTEM_DESC'],
                                                                                row['EGCRAFT'],
                                                                                row['PTW'])
                pm_parent_dict['STATUS'] = status
                pm_parent_dict['LOCATION'] = location
                pm_parent_dict['ROUTE'] = route
                pm_parent_dict['LEADTIME'] = leadtime
                pm_parent_dict['PMCOUNTER'] =''
                pm_parent_dict['WORKTYPE'] = worktype
                pm_parent_dict['EGMNTACTTYPE'] = egmntacttype
                pm_parent_dict['WOSTATUS'] = wostatus
                pm_parent_dict['EGCRAFT'] = '_'.join(group['EGCRAFT'].drop_duplicates().to_list())
                pm_parent_dict['RESPONSED BY'] = '_'.join(group['RESPONSED BY'].drop_duplicates().to_list())
                pm_parent_dict['PTW'] = '_'.join(group['PTW'].drop_duplicates().to_list())
                pm_parent_dict['LOTO'] = ''
                pm_parent_dict['EGPROJECTID'] = egprojectid
                pm_parent_dict['EGWBS']  = egwbs
                pm_parent_dict['FREQUENCY'] = frequency
                pm_parent_dict['FREQUNIT'] = frequnit
                pm_parent_dict['NEXTDATE'] = group['NEXTDATE'].min()
                pm_parent_dict['TARGSTARTTIME'] = group['TARGSTARTTIME'].min()
                pm_parent_dict['FINISH_DATE'] = group['FINISH_DATE'].max()
                pm_parent_dict['FINISH TIME'] = group['FINISH TIME'].max()
                pm_parent_dict['PARENT'] = ''
                pm_parent_dict['JPNUM'] = ''
                pm_parent_dict['MAIN_SYSTEM'] =row['MAIN_SYSTEM']
            #     pm_parent_dict['SUB_SYSTEM'] =row['SUB_SYSTEM']
            #     pm_parent_dict['EQUIPMENT'] =row['EQUIPMENT']
                pm_parent_dict['MAIN_SYSTEM_DESC'] =row['MAIN_SYSTEM_DESC']
            #     pm_parent_dict['SUB_SYSTEM_DESC'] =row['SUB_SYSTEM_DESC']
                pm_parent_dict['UNIT_TYPE'] =row['UNIT_TYPE']
                pm_parent_dict['TYPE'] =row['TYPE']
                ##################
                parent_df = pd.DataFrame(pm_parent_dict,index=np.arange(1))
                if loop_num < 10:
                    group.loc[:,'PARENT'] = 'Group-0{}-{}-{}-{}'.format(loop_num,row['MAIN_SYSTEM'],row['TYPE'],row['UNIT_TYPE'])
                else:
                    group.loc[:,'PARENT'] = 'Group-{}-{}-{}-{}'.format(loop_num,row['MAIN_SYSTEM'],row['TYPE'],row['UNIT_TYPE'])
                    
                #group
                #############################
                parent = pd.concat([parent_df, group])
                parent.loc[:,'GROUP'] = loop_num
                df_pm_plan3_more = pd.concat([df_pm_plan3_more,parent])

            df_pm_plan3_less = pd.DataFrame()
            loop_num = 0
            #df_xx = pm_master_df[['MAIN_SYSTEM','MAIN_SYSTEM_DESC','EGCRAFT','PTW','UNIT_TYPE','TYPE']].drop_duplicates().sort_values(by=['TYPE','UNIT_TYPE'],ascending=False)
            for index,row in df_yyy_cont_less.iterrows():
                loop_num +=1 
                #print(row['system_eq'])
                #################
                group = pm_master_df[(pm_master_df['MAIN_SYSTEM']==row['MAIN_SYSTEM']) &
                                    (pm_master_df['MAIN_SYSTEM_DESC']==row['MAIN_SYSTEM_DESC']) &
                                    (pm_master_df['EGCRAFT']==row['EGCRAFT'])&
                                    (pm_master_df['PTW']==row['PTW'])&
                                    (pm_master_df['UNIT_TYPE']==row['UNIT_TYPE']) &
                                    (pm_master_df['TYPE']==row['TYPE'])]
                df_pm_plan3_less = pd.concat([df_pm_plan3_less, group])

            df_pm_plan3_more_ME = df_pm_plan3_more[df_pm_plan3_more['TYPE']=='ME']
            df_pm_plan3_more_EE = df_pm_plan3_more[df_pm_plan3_more['TYPE']=='EE']
            df_pm_plan3_more_CV = df_pm_plan3_more[df_pm_plan3_more['TYPE']=='CV']
            df_pm_plan3_more_IC = df_pm_plan3_more[df_pm_plan3_more['TYPE']=='IC']

            df_pm_plan3_less_ME = df_pm_plan3_less[df_pm_plan3_less['TYPE']=='ME']
            df_pm_plan3_less_EE = df_pm_plan3_less[df_pm_plan3_less['TYPE']=='EE']
            df_pm_plan3_less_CV = df_pm_plan3_less[df_pm_plan3_less['TYPE']=='CV']
            df_pm_plan3_less_IC = df_pm_plan3_less[df_pm_plan3_less['TYPE']=='IC']

            lst_group  = [df_pm_plan3_more_ME,df_pm_plan3_more_EE,df_pm_plan3_more_CV,df_pm_plan3_more_IC]
            lst_no_group = [df_pm_plan3_less_ME,df_pm_plan3_less_EE, df_pm_plan3_less_CV,df_pm_plan3_less_IC]

            df_pm_plan3_test = pd.DataFrame()
            for i in range(len(lst_group)):
                df_pm_plan3_test = pd.concat([df_pm_plan3_test,lst_group[i]])
                df_pm_plan3_test = pd.concat([df_pm_plan3_test,lst_no_group[i]])

            df_pm_plan3_test['MOD'] = df_pm_plan3_test['GROUP']%2
            df_pm_plan3 = df_pm_plan3_test.copy()

            df_pm_plan3['PARENTCHGSSTATUS']=''
            df_pm_plan3['WOSEQUENCE']=''

            lst_new = ['PMNUM', 'SITEID', 'DESCRIPTION', 'STATUS','LOCATION','WORKTYPE',
                    'EGMNTACTTYPE','WOSTATUS','EGPROJECTID','EGWBS', 'FREQUENCY', 'FREQUNIT',
                    'JPNUM', 'ROUTE', 'NEXTDATE', 'EGCRAFT', 'RESPONSED BY', 'PTW', 'PARENTCHGSSTATUS',
                    'LEADTIME', 'TARGSTARTTIME', 'PARENT', 'PMCOUNTER', 'WOSEQUENCE',
                    'MOD','MAIN_SYSTEM','MAIN_SYSTEM_DESC','UNIT_TYPE','TYPE', 'SUB_SYSTEM', 'EQUIPMENT', 'SUB_SYSTEM_DESC',
                    'KKS_NEW_DESC', 'GROUP', 'FINISH_DATE', 'FINISH TIME']

            df_pm_plan3_master = df_pm_plan3[lst_new].reset_index(drop=True)
            index_parent = df_pm_plan3_master[(df_pm_plan3_master['PARENT']=='')].index
            index_child = df_pm_plan3_master[(df_pm_plan3_master['PARENT']!='')].index

            df_pm_plan3_master.loc[index_parent,'PARENTCHGSSTATUS']=0
            df_pm_plan3_master.loc[index_child,'PARENTCHGSSTATUS']=1
            df_pm_plan3_master.loc[:,'PMCOUNTER']=0

            index_parent2 = df_pm_plan3_master[df_pm_plan3_master['PARENTCHGSSTATUS'] == 0].index
            lst_parent_group = [i for i in range(1,len(index_parent2)+1)]
            df_pm_plan3_master.loc[index_parent2,'WOSEQUENCE'] = lst_parent_group

            for group in df_pm_plan3_master['PARENT'].value_counts().index:
                if group!='':
                    index_child = df_pm_plan3_master[df_pm_plan3_master['PARENT']==group].index
                    run_number = [i for i in range(1,len(index_child)+1)]
                    df_pm_plan3_master.loc[index_child,'WOSEQUENCE'] = run_number

            for key, value in user_input_mapping.items():
                if value.startswith('C'):
                    # Define condition
                    pm_plan_cond = ((df_pm_plan3_master['UNIT_TYPE'] == value)) & \
                                    (df_pm_plan3_master['PARENT'] == '') & \
                                    (df_pm_plan3_master['JPNUM'] == '')
                    
                    # Update LOCATION column
                    df_pm_plan3_master.loc[pm_plan_cond, 'LOCATION'] = f'{first_plant}{key}'

            df_pm_plan3_master['NEXTDATE'] = df_pm_plan3_master['NEXTDATE'].astype('str')
            df_pm_plan3_master['FINISH_DATE'] = df_pm_plan3_master['FINISH_DATE'].astype('str')
            df_pm_plan3_master['TARGSTARTTIME'] = df_pm_plan3_master['TARGSTARTTIME'].astype('str')
            df_pm_plan3_master['FINISH TIME'] = df_pm_plan3_master['FINISH TIME'].astype('str')

            df_pm_plan3_master['COMMENT'] = ''
            pm_plan_cond0 = df_pm_plan3_master['PMNUM'].str.len()>30
            # comment_empty_or_na = (df_pm_plan3_master['COMMENT'] == '') | (df_pm_plan3_master['COMMENT'].isna())
            # df_pm_plan3_master.loc[pm_plan_cond0 & comment_empty_or_na, 'COMMENT'] = 'PMNUM มีความยาวมากกว่า 30 ตัวอักษร'
            # df_pm_plan3_master.loc[pm_plan_cond0 & ~comment_empty_or_na, 'COMMENT'] += ', PMNUM มีความยาวมากกว่า 30 ตัวอักษร'
            update_comment(df_pm_plan3_master, pm_plan_cond0, 'COMMENT', 'PMNUM มีความยาวมากกว่า 30 ตัวอักษร')

            # Log ใช้ในการทำ Scheduler
            display_text_cond1 = df_pm_plan3_master['PARENT'] == ''
            display_text = ', '.join(df_pm_plan3_master[display_text_cond1]['PMNUM'].astype(str).tolist())
            print(display_text)

            #! Create PM_Plan.xlsx
            df_pm_plan3_master.to_excel(pm_plan_path, index=False)

            #! TYPE 2
            class PMNumGenerator:
                def __init__(self):
                    self.primary_counter = 0
                    self.secondary_counter = 0
                    self.sub_counter = 0
                    self.pmnum0 = None
                
                def create_pmnum(self, row):
                    self.primary_counter += 10
                    
                    try:
                        if row['PARENTCHGSSTATUS'] == 0 and (row['JPNUM'] == '' or pd.isna(row['JPNUM'])):
                            self.secondary_counter += 1
                            self.sub_counter = 0
                            self.pmnum0 = f"MI-{location}-{str(self.primary_counter).zfill(5)}-{row['TYPE']}{str(self.secondary_counter).zfill(2)}-{str(self.sub_counter).zfill(2)}"
                            pmnum1 = self.pmnum0
                            jpnum1 = ''
                            parent1 = ''
                        
                        elif row['PARENTCHGSSTATUS'] == 1 and pd.notna(row['JPNUM']):
                            self.sub_counter += 1
                            pmnum1 = f"MI-{location}-{str(self.primary_counter).zfill(5)}-{row['TYPE']}{str(self.secondary_counter).zfill(2)}-{str(self.sub_counter).zfill(2)}"
                            jpnum1 = f"JP-{pmnum1}"
                            parent1 = self.pmnum0
                        
                        else:
                            self.secondary_counter += 1
                            pmnum1 = f"MI-{location}-{str(self.primary_counter).zfill(5)}-{row['TYPE']}{str(self.secondary_counter).zfill(2)}"
                            jpnum1 = f"JP-{pmnum1}"
                            parent1 = ''
                    
                    except KeyError as e:
                        logger.error(f"Column not found: {e}")
                    except Exception as e:
                        logger.error(f"Error processing row: {e}")
                    
                    return pd.Series([self.pmnum0, pmnum1, jpnum1, parent1])

            # Instantiate the class
            generator = PMNumGenerator()

            # Apply the function
            df_pm_plan3_master[['PMNUM1', 'PMNUM1', 'JPNUM1', 'PARENT1']] = df_pm_plan3_master.apply(generator.create_pmnum, axis=1)
            
            jpnum_map = df_pm_plan3_master.set_index('JPNUM')['JPNUM1'].to_dict()
            df_jop_plan_master['JOB_NUM1'] = df_jop_plan_master['JOB_NUM'].map(jpnum_map)
            df_labor_new['JOB_NUM1'] = df_jop_plan_master['JOB_NUM'].map(jpnum_map)

            df_pm_plan3_master['COMMENT1'] = ''
            pm_plan_cond1 = df_pm_plan3_master['PMNUM1'].str.len()>30
            update_comment(df_pm_plan3_master, pm_plan_cond1, 'COMMENT1', 'PMNUM มีความยาวมากกว่า 30 ตัวอักษร')

            df_jop_plan_master['COMMENT1'] = ''
            job_plan_cond1 = df_jop_plan_master['JOB_NUM1'].str.len()>30
            update_comment(df_jop_plan_master, job_plan_cond1, 'COMMENT1', 'JPNUM มีความยาวมากกว่า 30 ตัวอักษร')

            df_labor_new['COMMENT1'] = ''
            job_plan_cond1 = df_labor_new['JOB_NUM1'].str.len()>30
            update_comment(df_labor_new, job_plan_cond1, 'COMMENT1', 'JPNUM มีความยาวมากกว่า 30 ตัวอักษร')
            
            pm_plan_column1 = ['PMNUM', 'SITEID', 'DESCRIPTION', 'STATUS', 'LOCATION', 'WORKTYPE',
                                    'EGMNTACTTYPE', 'WOSTATUS', 'EGPROJECTID', 'EGWBS', 'FREQUENCY',
                                    'FREQUNIT', 'JPNUM', 'ROUTE', 'NEXTDATE', 'EGCRAFT', 'RESPONSED BY',
                                    'PTW', 'PARENTCHGSSTATUS', 'LEADTIME', 'TARGSTARTTIME', 'PARENT',
                                    'PMCOUNTER', 'WOSEQUENCE', 'MOD', 'MAIN_SYSTEM', 'MAIN_SYSTEM_DESC',
                                    'UNIT_TYPE', 'TYPE', 'SUB_SYSTEM', 'EQUIPMENT', 'SUB_SYSTEM_DESC',
                                    'KKS_NEW_DESC', 'GROUP', 'FINISH_DATE', 'FINISH TIME', 'COMMENT']
            replace_dict = {
                'PMNUM': 'PMNUM1',
                'JPNUM': 'JPNUM1',
                'PARENT': 'PARENT1',
                'COMMENT': 'COMMENT1'
            }
            pm_plan_column2 = [replace_columns(col, replace_dict) for col in pm_plan_column1]


            jop_plan_column1 = ['JOB_NUM', 'ORGID', 'SITEID', 'PLUSCREVNUM', 'STATUS', 'EQUIPMENT', 
                                    'DURATION_TOTAL', 'TASK_ORDER_NEW', 'PLUSCJPREVNUM', 'DURATION_(HR.)', 
                                    'TASK_NEW', 'START_DATE', 'FINISH_DATE', 'GROUP', 'MOD', 'COMMENT']
            replace_dict = {
                'JOB_NUM': 'JOB_NUM1',
                'COMMENT': 'COMMENT1',
            }
            jop_plan_column2 = [replace_columns(col, replace_dict) for col in jop_plan_column1]

            lst_labor2 = [replace_columns(col, replace_dict) for col in lst_labor1]
            
            #! Create WORKORDER
            
            # ใช้ฟังก์ชันสำหรับ PMNUM
            df_workorder1 = create_workorder(df_pm_plan3_master, 'PMNUM', orgid)
            # ใช้ฟังก์ชันสำหรับ PMNUM1
            df_workorder2 = create_workorder(df_pm_plan3_master, 'PMNUM1', orgid)
            
            #! Create Template-MxLoader-JP-PMPlan
            # Define the color and border style
            color = "666666"
            thin_border = Border(
                left=Side(style='thin', color=color),
                right=Side(style='thin', color=color),
                top=Side(style='thin', color=color),
                bottom=Side(style='thin', color=color)
            )

            # Define the background color fill
            fill_color = PatternFill(start_color='F5F5DC', end_color='F5F5DC', fill_type='solid')
            yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')

            # Define file paths
            file_path = os.path.join(settings.STATIC_ROOT, 'excel', 'Template-MxLoader-JP-PMPlan.xlsm')
            file_template_xlsx = os.path.join(temp_dir, f"{uuid.uuid4()}_Template-MxLoader-JP-PMPlan_temp.xlsx")
            file_template_xlsm = os.path.join(temp_dir, f"{uuid.uuid4()}_Template-MxLoader-JP-PMPlan.xlsm")
            
            sheet_jp_labor = 'JPPLAN-LABOR'
            sheet_jp_task = 'JPPLAN-TASK'
            sheet_pm = 'PMPlan'
            sheet_wo = 'WO'
            start_row = 2

            # 1. ตรวจสอบว่ามีไฟล์ที่ต้องการใช้งานอยู่จริงหรือไม่
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"The file {file_path} does not exist.")

            # 2. คัดลอกไฟล์ต้นฉบับไปยังไฟล์ใหม่ (file_template_xlsx)
            shutil.copyfile(file_path, file_template_xlsx)

            try:
                basic_sheet_names = [sheet_jp_labor, sheet_jp_task, sheet_pm, sheet_wo]
                for i in range(1, 2):
                    for sheet_name in basic_sheet_names:
                        copy_worksheet(file_template_xlsx, sheet_name, i)
                
                # 3. ใช้ pandas เพื่อเขียนข้อมูลลงในไฟล์ .xlsx
                # with pd.ExcelWriter(file_template_xlsx, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                # # เขียนข้อมูลจาก DataFrame ลงในชีตที่ระบุ
                #     df_labor_new[lst_labor1].to_excel(writer, sheet_name=sheet_jp_labor, startrow=start_row, startcol=0, index=False, header=False)
                #     df_jop_plan_master[jop_plan_column1].to_excel(writer, sheet_name=sheet_jp_task, startrow=start_row, startcol=0, index=False, header=False)
                #     df_pm_plan3_master[pm_plan_column1].to_excel(writer, sheet_name=sheet_pm, startrow=start_row, startcol=0, index=False, header=False)
                #     df_workorder1.to_excel(writer, sheet_name=sheet_wo, startrow=(start_row+5), startcol=1, index=False, header=False)
                
                with pd.ExcelWriter(file_template_xlsx, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                    # ชุดที่ 1
                    write_dataframes_to_excel(
                        writer, 
                        df_labor_new, lst_labor1, 
                        df_jop_plan_master, jop_plan_column1, 
                        df_pm_plan3_master, pm_plan_column1, 
                        df_workorder1, 
                        sheet_jp_labor, sheet_jp_task, sheet_pm, sheet_wo, 
                        start_row, start_offset=5
                    )
                    
                    # ชุดที่ 2
                    write_dataframes_to_excel(
                        writer, 
                        df_labor_new, lst_labor2, 
                        df_jop_plan_master, jop_plan_column2, 
                        df_pm_plan3_master, pm_plan_column2, 
                        df_workorder2, 
                        f"{sheet_jp_labor}-1", f"{sheet_jp_task}-1", f"{sheet_pm}-1", f"{sheet_wo}-1", 
                        start_row, start_offset=5
                    )

                openpyxl_book = load_workbook(file_template_xlsx)
                
                for sheet_name in basic_sheet_names:
                    sheet = openpyxl_book[sheet_name]
                    decorate_sheet(sheet, sheet_name, thin_border, fill_color, yellow_fill, start_row, df_jop_plan_master, df_pm_plan3_master)

                # วนลูปสำหรับชีตที่มี -1 ต่อท้าย
                for sheet_name in basic_sheet_names:
                    sheet_name_with_suffix = f"{sheet_name}-1"
                    sheet = openpyxl_book[sheet_name_with_suffix]
                    decorate_sheet(sheet, sheet_name_with_suffix, thin_border, fill_color, yellow_fill, start_row, df_jop_plan_master, df_pm_plan3_master)
                
                # บันทึกไฟล์ตกแต่งในรูปแบบ .xlsx
                openpyxl_book.save(file_template_xlsx)
                
                # 5. ใช้ xlwings คัดลอกเนื้อหาและการตกแต่งจาก .xlsx ไปยัง .xlsm
                with xw.App(visible=False) as app:
                    time.sleep(2)

                    # เปิดไฟล์ .xlsx และ .xlsm
                    book_xlsx = app.books.open(file_template_xlsx)
                    book_xlsm = app.books.open(file_path)  # เปิดไฟล์ต้นฉบับ

                    # ลบชีตเดิมใน .xlsm ที่ต้องการแทนที่
                    for sheet_name in basic_sheet_names:
                        if sheet_name in [s.name for s in book_xlsm.sheets]:
                            book_xlsm.sheets[sheet_name].delete()
                    
                    # คัดลอกชีตทั้งหมดจาก .xlsx ไปยัง .xlsm
                    for sheet_name in basic_sheet_names:
                        sheet_xlsx = book_xlsx.sheets[sheet_name]
                        # คัดลอกชีตจาก .xlsx ไปยัง .xlsm วางไว้หลังชีตสุดท้ายใน .xlsm
                        sheet_xlsx.api.Copy(After=book_xlsm.sheets[-1].api)

                        # ตั้งชื่อชีตใหม่ใน .xlsm
                        new_sheet = book_xlsm.sheets[-1]  # ชีตที่ถูกคัดลอกจะเป็นชีตสุดท้าย
                        if sheet_name in [s.name for s in book_xlsm.sheets]:
                            new_sheet_name = f"{sheet_name}-{location}"
                        else:
                            new_sheet_name = sheet_name

                        new_sheet.name = new_sheet_name
                        
                    for sheet_name in basic_sheet_names:
                        sheet_name_with_suffix = f"{sheet_name}-1"
                        sheet_xlsx = book_xlsx.sheets[sheet_name_with_suffix]
                        # คัดลอกชีตจาก .xlsx ไปยัง .xlsm วางไว้หลังชีตสุดท้ายใน .xlsm
                        sheet_xlsx.api.Copy(After=book_xlsm.sheets[-1].api)

                        # ตั้งชื่อชีตใหม่ใน .xlsm
                        new_sheet = book_xlsm.sheets[-1]  # ชีตที่ถูกคัดลอกจะเป็นชีตสุดท้าย
                        if sheet_name_with_suffix in [s.name for s in book_xlsm.sheets]:
                            new_sheet_name = f"{sheet_name_with_suffix}-{location}"
                        else:
                            new_sheet_name = sheet_name_with_suffix

                        new_sheet.name = new_sheet_name

                    # บันทึกไฟล์ .xlsm ที่ตกแต่งแล้ว
                    book_xlsm.save(file_template_xlsm)
                    logger.info(f"File saved successfully as {file_template_xlsm}")

            except Exception as e:
                print(f"An error occurred: {e}")

            finally:
                # ปิดไฟล์
                try:
                    openpyxl_book.close()
                except Exception as e:
                    logger.error(f"Error closing the openpyxl book: {e}")
                
                try:
                    os.remove(file_template_xlsx)
                except FileNotFoundError:
                    logger.warning(f"Temporary file {file_template_xlsx} not found for deletion.")
                except Exception as e:
                    logger.error(f"Error deleting temporary file: {e}")
############
############
            # บันทึกลิงก์ดาวน์โหลดลงใน session
            request.session['download_link_comment'] = comment_path
            request.session['download_link_job_plan_task'] = job_plan_task_path
            request.session['download_link_job_plan_labor'] = job_plan_labor_path
            request.session['download_link_pm_plan'] = pm_plan_path
            request.session['download_link_template'] = file_template_xlsm
            
            #! Check session
            print('form: ',form)
            print('schedule_filename:',schedule_filename)
            print('location_filename:',location_filename)
            print('extracted_kks_counts:',extracted_kks_counts)
            print('user_input_mapping:',user_input_mapping) 
            print('download_link_comment:', request.session['download_link_comment'])
            print('download_link_job_plan_task:',  request.session['download_link_job_plan_task'])
            print('download_link_job_plan_labor:',  request.session['download_link_job_plan_labor'])
            print('download_link_pm_plan:',  request.session['download_link_pm_plan'])
            print('download_link_template:',  request.session['download_link_template'])
        
        else:
            # ถ้าฟอร์มไม่ถูกต้อง ให้แสดงฟอร์มพร้อมกับข้อมูลเดิม
            return render(request, 'maximo_app/upload.html', {
                'form': form,
                'error_message': error_message,
            })
    else:
        # เมื่อไม่มีการส่งข้อมูลผ่านแบบฟอร์ม (เช่น เมื่อผู้ใช้เข้าถึงหน้าเว็บครั้งแรกหรือรีเฟรชหน้าเว็บโดยไม่มีการส่งแบบฟอร์ม) 
        # จะมีการสร้างอินสแตนซ์ของ UploadFileForm() ซึ่งเป็นฟอร์มที่ใช้ในการอัปโหลดไฟล์ เพื่อให้ฟอร์มแสดงในหน้าเว็บ 
        # ผู้ใช้สามารถเห็นฟอร์มและอัปโหลดไฟล์ได้
        
        # #! Clear session data after download
        # request.session.pop('schedule_filename', None)
        # request.session.pop('location_filename', None)
        # request.session.pop('extracted_kks_counts', None)
        # request.session.pop('first_plant', None)
        # request.session.pop('most_common_plant_unit', None)
        # request.session.pop('df_original', None)
        # request.session.pop('df_original_copy', None)
        # request.session.pop('df_original_newcol', None)
        # request.session.pop('df_comment', None)
        # request.session.pop('user_input_mapping', None)
        # request.session.pop('download_link_comment', None)
        # request.session.pop('download_link_job_plan_task', None)
        # request.session.pop('download_link_job_plan_labor', None)
        # request.session.pop('download_link_pm_plan', None)
        # request.session.pop('download_link_template', None)
        
        # #! Dropdown
        # request.session.pop('egmntacttype', None)
        # request.session.pop('egprojectid', None)
        # request.session.pop('egwbs', None)
        # request.session.pop('location', None)
        # request.session.pop('siteid', None)
        # request.session.pop('wbs_desc', None)
        # request.session.pop('worktype', None)
        # request.session.pop('wostatus', None)
        
        request.session.clear()
        form = UploadFileForm()
    
    return render(request, 'maximo_app/upload.html', {
        'form': form,
        'error_message': error_message,
        'schedule_filename': schedule_filename,
        'location_filename': location_filename,
        'extracted_kks_counts': extracted_kks_counts,
        'user_input_mapping' : user_input_mapping,
    })

# ---------------------------------
# ฟังก์ชันการจัดการการดาวน์โหลด (Download Functions)
# ---------------------------------

def download_comment_file(request):
    # ดึง path ของไฟล์จาก session
    comment_file = request.session.get('download_link_comment', None)
    original_file_name = 'Comment.xlsx'

    # ตรวจสอบว่ามีการตั้งค่า comment_file หรือไม่
    if comment_file:
        # ตรวจสอบว่ามีการอ้างอิงไฟล์จริงใน path ที่ถูกต้องหรือไม่
        full_comment_file_path = os.path.abspath(comment_file)
        if os.path.exists(full_comment_file_path):
            try:
                # เปิดไฟล์และสร้าง response สำหรับการดาวน์โหลด
                with open(full_comment_file_path, 'rb') as fh:
                    response = HttpResponse(fh.read(), content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    response['Content-Disposition'] = f'attachment; filename={original_file_name}'
                    return response
            except IOError:
                # แสดงข้อผิดพลาดเมื่อไม่สามารถอ่านไฟล์ได้
                raise Http404("Error reading file")
        else:
            # แสดงข้อผิดพลาดเมื่อไม่พบไฟล์
            raise Http404("File not found")
    else:
        # แสดงข้อผิดพลาดเมื่อไม่ได้รับ path ของไฟล์จาก session
        raise Http404("No file specified for download")

def download_job_plan_task_file(request):
    # ดึงลิงก์ไฟล์จาก session
    jp_task_file = request.session.get('download_link_job_plan_task', None)
    original_file_name = 'Job_Plan_Task.xlsx'

    if jp_task_file:
        # ตรวจสอบเส้นทางของไฟล์
        full_jp_file_path = os.path.abspath(jp_task_file)
        
        # ตรวจสอบว่าไฟล์มีอยู่จริงหรือไม่
        if os.path.exists(full_jp_file_path):
            try:
                with open(full_jp_file_path, 'rb') as fh:
                    response = HttpResponse(fh.read(), content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    response['Content-Disposition'] = f'attachment; filename={original_file_name}'
                    return response
            except IOError:
                raise Http404("Error reading file")
        else:
            raise Http404("File not found")
    else:
        raise Http404("No file specified for download")

def download_job_plan_labor_file(request):
    # ดึงลิงก์ไฟล์จาก session
    jp_labor_file = request.session.get('download_link_job_plan_labor', None)
    original_file_name = 'Job_Plan_Labor.xlsx'

    if jp_labor_file:
        # ตรวจสอบเส้นทางไฟล์
        full_jp_labor_file_path = os.path.abspath(jp_labor_file)
        
        # ตรวจสอบว่าไฟล์มีอยู่จริง
        if os.path.exists(full_jp_labor_file_path):
            try:
                # ใช้ with block เพื่อจัดการไฟล์อย่างปลอดภัย
                with open(full_jp_labor_file_path, 'rb') as fh:
                    response = HttpResponse(fh.read(), content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    response['Content-Disposition'] = f'attachment; filename={original_file_name}'
                    return response
            except IOError:
                raise Http404("Error reading file")
        else:
            raise Http404("File not found")
    else:
        raise Http404("No file specified for download")

def download_pm_plan_file(request):
    # ดึงลิงก์ไฟล์จาก session
    pm_plan_file = request.session.get('download_link_pm_plan', None)
    original_file_name = 'PM_Plan.xlsx'

    if pm_plan_file:
        # ตรวจสอบว่าเส้นทางของไฟล์เป็นเส้นทางสมบูรณ์
        full_pm_plan_file_path = os.path.abspath(pm_plan_file)
        
        # ตรวจสอบว่าไฟล์มีอยู่จริง
        if os.path.exists(full_pm_plan_file_path):
            try:
                # ใช้ with block เพื่อจัดการการเปิดไฟล์อย่างปลอดภัย
                with open(full_pm_plan_file_path, 'rb') as fh:
                    response = HttpResponse(fh.read(), content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    response['Content-Disposition'] = f'attachment; filename={original_file_name}'
                    return response
            except IOError:
                raise Http404("Error reading file")
        else:
            raise Http404("File not found")
    else:
        raise Http404("No file specified for download")

def download_template_file(request):
    # ดึงลิงก์ไฟล์จาก session
    template_file = request.session.get('download_link_template', None)
    original_file_name = 'Template-MxLoader-JP-PMPlan_(Blank).xlsm'

    if template_file:
        # ตรวจสอบเส้นทางของไฟล์ให้เป็นเส้นทางสมบูรณ์
        full_template_file_path = os.path.abspath(template_file)
        
        # ตรวจสอบว่าไฟล์มีอยู่จริง
        if os.path.exists(full_template_file_path):
            try:
                # เปิดไฟล์อย่างปลอดภัยและอ่านข้อมูล
                with open(full_template_file_path, 'rb') as fh:
                    response = HttpResponse(fh.read(), content_type="application/vnd.ms-excel.sheet.macroEnabled.12")
                    response['Content-Disposition'] = f'attachment; filename={original_file_name}'
                    return response
            except IOError:
                raise Http404("Error reading file")
        else:
            raise Http404("File not found")
    else:
        raise Http404("No file specified for download")

# ---------------------------------
# ฟังก์ชันการกรองข้อมูล (Filter Functions)
# ---------------------------------

@require_GET
def filter_site(request):
    site_id = request.GET.get('site_id')  # รับค่า site_id ที่ถูกส่งมาจาก request
    
    if not site_id:
        return JsonResponse({'error': 'No site_id provided.'}, status=400)
    
    try:
        site_id = int(site_id)
    except ValueError:
        return JsonResponse({'error': 'Invalid site_id format. Must be an integer.'}, status=400)
    
    try:
        site = Site.objects.get(id=site_id) # ดึงข้อมูล Site จากฐานข้อมูลโดยใช้ site_id
        logger.info(f"Successfully retrieved Site with id {site_id}")
        child_sites = site.child_sites.values('id', 'site_id', 'site_name')
        
    except Site.DoesNotExist:
        logger.warning(f"Site with id {site_id} does not exist.")
        return JsonResponse({'error': 'Site not found.'}, status=404)
    except Exception as e:
        logger.error(f"Unexpected error in filter_site: {str(e)}", exc_info=True)
        return JsonResponse({'error': 'An unexpected error occurred.'}, status=500)

    child_site_list = [{
        'id': child_site['id'],
        'site_id': child_site['site_id'],
        'site_name': child_site['site_name']
    } for child_site in child_sites]
    
    return JsonResponse({
        'site_name': site.site_name,
        'child_sites': child_site_list,
    })

@require_GET
def filter_child_site(request):
    # ดึงค่า child_site_id จาก request GET
    child_site_id = request.GET.get('child_site_id')

    if not child_site_id:
        return JsonResponse({'error': 'No child_site_id provided.'}, status=400)
    
    try:
        child_site = ChildSite.objects.get(id=child_site_id)
        logger.info(f"Successfully retrieved ChildSite with id {child_site_id}")
    except ChildSite.DoesNotExist:
        logger.warning(f"ChildSite with id {child_site_id} does not exist.")
        return JsonResponse({'error': 'Child Site not found.'}, status=404)
    except Exception as e:
        logger.error(f"Unexpected error in filter_child_site: {str(e)}")
        return JsonResponse({'error': 'An unexpected error occurred.'}, status=500)

    return JsonResponse({
        'description': child_site.site_name
    })

@require_GET
def filter_worktype(request):
    work_type_id = request.GET.get('work_type_id')  # รับค่า work_type_id จาก request
    
    if not work_type_id:
        return JsonResponse({'error': 'No work_type_id provided.'}, status=400)
    
    try:
        work_type_id = int(work_type_id)
    except ValueError:
        return JsonResponse({'error': 'Invalid work_type_id format. Must be an integer.'}, status=400)
    
    try:
        work_type = WorkType.objects.get(id=work_type_id)
        logger.info(f"Successfully retrieved WorkType with id {work_type_id}")
    except WorkType.DoesNotExist:
        logger.warning(f"WorkType with id {work_type_id} does not exist.")
        return JsonResponse({'error': 'WorkType not found.'}, status=404)
    except Exception as e:
        logger.error(f"Unexpected error in filter_worktype: {str(e)}", exc_info=True)
        return JsonResponse({'error': 'An unexpected error occurred.'}, status=500)
    return JsonResponse({'description': work_type.description})

@require_GET
def filter_plant_type(request):
    plant_type_id = request.GET.get('plant_type_id')
    
    if not plant_type_id:
        return JsonResponse({'error': 'No plant_type_id provided.'}, status=400)
    
    try:
        plant_type_id = int(plant_type_id)
    except ValueError:
        return JsonResponse({'error': 'Invalid plant_type_id format. Must be an integer.'}, status=400)
    
    try:
        # ดึงข้อมูล PlantType ตาม id
        plant_type = PlantType.objects.get(id=plant_type_id)
        logger.info(f"Successfully retrieved PlantType with id {plant_type_id}")
        # Get the ActTypes associated with the PlantType
        acttypes = plant_type.act_types.all()
        
        # Get the Sites associated with the PlantType
        sites = plant_type.sites.all()
        
        # Get the WorkType associated with the PlantType
        work_types = plant_type.work_types.all()
        
        # Get the Units associated with the PlantType
        units = plant_type.units.all()
    except PlantType.DoesNotExist:
        logger.warning(f"PlantType with id {plant_type_id} does not exist.")
        return JsonResponse({'error': 'PlantType not found.'}, status=404)
    except Exception as e:
        logger.error(f"Unexpected error in filter_plant_type: {str(e)}", exc_info=True)
        return JsonResponse({'error': 'An unexpected error occurred.'}, status=500)
    
    # แปลงข้อมูลเป็น list เพื่อส่งกลับไปยัง frontend
    acttype_list = [{'id': acttype.id, 'acttype': acttype.acttype} for acttype in acttypes]
    site_list = [{'id': site.id, 'site_id': site.site_id, 'site_name': site.site_name} for site in sites]    
    work_type_list = [{'id': worktype.id, 'worktype': worktype.worktype} for worktype in work_types]
    unit_list = [{'id': unit.id, 'unit_code': unit.unit_code} for unit in units]
    return JsonResponse({
        'acttypes': acttype_list,
        'sites': site_list,
        'plant_type_th': plant_type.plant_type_th,
        'work_types': work_type_list,
        'units' : unit_list,
    })
    
@require_GET
def filter_acttype(request):
    acttype_id = request.GET.get('acttype_id')

    if not acttype_id:
        return JsonResponse({'error': 'No acttype_id provided.'}, status=400)

    try:
        acttype_id = int(acttype_id)
    except ValueError:
        return JsonResponse({'error': 'Invalid acttype_id format. Must be an integer.'}, status=400)

    try:
        acttype = ActType.objects.get(id=acttype_id)
        logger.info(f"Successfully retrieved ActType with id {acttype_id}")
    except ActType.DoesNotExist:
        logger.warning(f"ActType with id {acttype_id} does not exist.")
        return JsonResponse({'error': 'ActType not found.'}, status=404)
    except Exception as e:
        logger.error(f"Unexpected error in filter_acttype: {str(e)}", exc_info=True)
        return JsonResponse({'error': 'An unexpected error occurred.'}, status=500)
    
    return JsonResponse({'description': acttype.description, 'code': acttype.code})

@require_GET
def filter_wbs(request):
    wbs_id = request.GET.get('wbs_id')

    if not wbs_id:
        return JsonResponse({'error': 'No wbs_id provided.'}, status=400)

    try:
        wbs_id = int(wbs_id)
    except ValueError:
        return JsonResponse({'error': 'Invalid wbs_id format. Must be an integer.'}, status=400)

    try:
        wbs = WBSCode.objects.get(id=wbs_id)
        logger.info(f"Successfully retrieved WBSCode with id {wbs_id}")
    except WBSCode.DoesNotExist:
        logger.warning(f"WBSCode with id {wbs_id} does not exist.")
        return JsonResponse({'error': 'WBSCode not found.'}, status=404)
    except Exception as e:
        logger.error(f"Unexpected error in filter_wbs: {e}", exc_info=True)
        return JsonResponse({'error': 'An unexpected error occurred.'}, status=500)
    
    return JsonResponse({'description': wbs.description})

@require_GET
def filter_wostatus(request):
    wostatus_id = request.GET.get('wostatus_id')

    if not wostatus_id:
        return JsonResponse({'error': 'No wostatus_id provided.'}, status=400)

    try:
        wostatus_id = int(wostatus_id)
    except ValueError:
        return JsonResponse({'error': 'Invalid wostatus_id format. Must be an integer.'}, status=400)

    try:
        wostatus = Status.objects.get(id=wostatus_id)
        logger.info(f"Successfully retrieved Status with id {wostatus_id}")
    except Status.DoesNotExist:
        logger.warning(f"Status with id {wostatus_id} does not exist.")
        return JsonResponse({'error': 'Status not found.'}, status=404)
    except Exception as e:
        logger.error(f"Unexpected error in filter_wostatus: {e}", exc_info=True)
        return JsonResponse({'error': 'An unexpected error occurred.'}, status=500)
    
    return JsonResponse({'description': wostatus.description})

# ---------------------------------
# ส่วนของ Custom Error Handlers
# ---------------------------------

def custom_404(request, exception):
    return render(request, '404.html', status=404)

def custom_500(request):
    return render(request, '500.html', status=500)

# ---------------------------------
# ฟังก์ชันช่วยเหลือ (Helper Functions)
# ---------------------------------

# ฟังก์ชันสำหรับแปลงคอลัมน์ที่เป็น Timestamp เป็น string
def convert_timestamp_columns_to_str(df, columns):
    for col in columns:
        if col in df.columns:
            df[col] = df[col].astype(str)
    return df

# ฟังก์ชันสำหรับแปลงคอลัมน์ string กลับเป็น Timestamp
def convert_str_columns_to_timestamp(df, columns):
    for col in columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')  # แปลงเป็น Timestamp
    return df

def update_comment(df, condition, column_name, message):
    comment_empty_or_na = (df[column_name] == '') | (df[column_name].isna())
    df.loc[condition & comment_empty_or_na, column_name] = message
    df.loc[condition & ~comment_empty_or_na, column_name] += f", {message}"

def replace_or_append_comment(df, condition, comment_col, message, replace_message=None):
    # กรณีที่ COMMENT ว่างหรือเป็น NaN ให้เขียนข้อความใหม่
    comment_empty_or_na = (df[comment_col] == '') | (df[comment_col].isna())
    df.loc[condition & comment_empty_or_na, comment_col] = message

    # กรณีที่ COMMENT มีข้อความแล้ว ให้เพิ่มหรือแทนที่ข้อความ
    df.loc[condition & ~comment_empty_or_na, comment_col] = df.loc[condition & ~comment_empty_or_na, comment_col].apply(
        lambda x: x.replace(replace_message, message) if replace_message in x else f"{x}, {message}"
    )

def read_excel_with_error_handling(schedule_path, sheet_name=0):
    try:
        # พยายามอ่านไฟล์ Excel ด้วย sheet_name ที่ระบุ
        df_original = pd.read_excel(schedule_path, 
                                    sheet_name=sheet_name, 
                                    header=1,
                                    dtype={'START_DATE': str, 'FINISH_DATE': str})
        return df_original
    except ValueError as ve:
        # หาก sheet_name ไม่ถูกต้อง
        logger.error(f"ข้อผิดพลาด: {ve}. กรุณาตรวจสอบว่าชื่อชีท '{sheet_name}' ถูกต้องหรือไม่.")
        return None
    except Exception as e:
        logger.error(f"ข้อผิดพลาด: {e}")
        return None

def convert_duration(value):
    try:
        value = float(value)
        # ตรวจสอบว่าเป็นจำนวนเต็มหรือไม่ หากใช่ ให้แปลงเป็น int
        if value.is_integer():
            return int(value)
        else:
            return value
    except (ValueError, TypeError):
        # หากไม่สามารถแปลงเป็น float ได้ หรือค่าที่ไม่สามารถแปลงได้
        return value

# ฟังก์ชันเพื่อตรวจสอบว่าค่าคือวันที่ตามรูปแบบ "DD-MMM-YYYY" หรือไม่
def is_date(value):
    try:
        # ตรวจสอบว่าค่าคือวันที่ในรูปแบบ "DD-MMM-YYYY"
        return bool(re.match(r"\d{1,2}-[A-Za-z]{3}-\d{4}", value))
    except TypeError:
        return False

def parse_dates(date_series):
    # แปลงวันที่ตามรูปแบบต่าง ๆ
    date_1 = pd.to_datetime(date_series, format='%d/%m/%Y', errors='coerce')
    date_2 = pd.to_datetime(date_series, format='%Y-%m-%d %H:%M:%S', errors='coerce')
    date_3 = pd.to_datetime(date_series, format='%d-%b-%Y', errors='coerce')
    date_4 = pd.to_datetime(date_series, format='%m/%d/%Y', errors='coerce')

    # รวมผลลัพธ์จากรูปแบบต่าง ๆ (ใช้ค่าแรกที่ไม่เป็น NaN)
    return date_1.fillna(date_2).fillna(date_3).fillna(date_4)

# Clean '//'
def clean_ptw_column(ptw_value):
    try:
        # ตรวจสอบว่าข้อมูลไม่เป็นค่าว่างหรือ None
        if not ptw_value or not isinstance(ptw_value, str):
            raise ValueError("ค่า PTW ต้องเป็นสตริงที่ไม่ว่าง")

        # แยกข้อความตาม '//'
        if '//' in ptw_value:
            # แยกข้อความตาม '//' ทั้งหมด
            parts = ptw_value.split('//')
            
            # ส่วนซ้ายจะเป็นส่วนแรกเสมอ
            left_side = parts[0]
            
            # ฝั่งขวาทั้งหมดรวมเข้าด้วยกัน
            right_side = ','.join([part.strip() for part in parts[1:]])

            # แปลงฝั่งซ้ายเป็น list เพื่อคงลำดับ
            left_side_list = [item.strip() for item in left_side.split(',')]
            # แปลงฝั่งขวาเป็น set เพื่อกำจัดคำซ้ำ
            right_side_set = set([item.strip() for item in right_side.split(',')])
            
            # ลบคำซ้ำจากฝั่งขวาที่มีในฝั่งซ้าย
            cleaned_right_side = [item for item in right_side_set if item not in left_side_list]
            
            # รวมค่าใหม่กลับมา
            cleaned_left_str = ', '.join(left_side_list)  # ฝั่งซ้ายคงอยู่เหมือนเดิม
            cleaned_right_str = ', '.join(cleaned_right_side)  # ลบคำซ้ำจากฝั่งขวา
            
            # ถ้าฝั่งขวามีค่า ให้รวมฝั่งซ้ายและขวาเข้าด้วยกัน
            if cleaned_right_str:
                return f'{cleaned_left_str}//{cleaned_right_str}'
            else:
                return cleaned_left_str  # ถ้าฝั่งขวาว่าง ให้แสดงเฉพาะฝั่งซ้าย
        else:
            return ptw_value  # ถ้าไม่มี '//' ในข้อความให้คืนค่าตามเดิม

    except Exception as e:
        logger.error(f"An error occurred: {e}")
        return ptw_value

def replace_columns(col, replace_dict):
    try:
        for key, value in replace_dict.items():
            col = re.sub(rf'\b{key}\b', value, col)
    except Exception as e:
        logger.error(f"An error occurred: {e}")
    return col

def create_workorder(df, pmnum_col, orgid):
    # เลือกคอลัมน์ที่ต้องการจาก DataFrame
    df_workorder = df[['SITEID', pmnum_col, 'WOSTATUS', 'NEXTDATE', 'FINISH_DATE']].copy()
    
    # กำหนดค่าใหม่ในคอลัมน์ต่างๆ
    df_workorder['PARENTCHGSSTATUS'] = 0
    df_workorder['ORGID'] = orgid
    df_workorder['NEXTDATE'] = df_workorder['NEXTDATE'] + ' ' + '01:00:00'
    df_workorder['FINISH_DATE'] = df_workorder['FINISH_DATE'] + ' ' + '09:00:00'
    df_workorder['SCHEDSTART'] = df_workorder['NEXTDATE']
    df_workorder['SCHEDFINISH'] = df_workorder['FINISH_DATE']
    df_workorder.rename(columns={'WOSTATUS': 'STATUS', 'NEXTDATE': 'TARGSTARTDATE', 'FINISH_DATE': 'TARGCOMPDATE'}, inplace=True)
    columns_order = ['SITEID', 'ORGID', 'STATUS', pmnum_col, 'PARENTCHGSSTATUS', 'TARGSTARTDATE', 'TARGCOMPDATE', 'SCHEDSTART', 'SCHEDFINISH']
    df_workorder = df_workorder[columns_order]
    
    return df_workorder

def copy_worksheet(file_path, sheet_name, index):
    wb = load_workbook(file_path)
    source_sheet = wb[sheet_name]
    new_sheet = wb.copy_worksheet(source_sheet)
    new_sheet.title = f"{sheet_name}-{index}"
    wb.save(file_path)

def write_dataframes_to_excel(writer, df_labor, lst_labor, df_jop, jop_plan_column, df_pm, pm_plan_column, df_workorder, sheet_name_labor, sheet_name_task, sheet_name_pm, sheet_name_wo, start_row, start_offset):
    # เขียนข้อมูลจาก DataFrame ลงในชีตที่ระบุ
    df_labor[lst_labor].to_excel(writer, sheet_name=sheet_name_labor, startrow=start_row, startcol=0, index=False, header=False)
    df_jop[jop_plan_column].to_excel(writer, sheet_name=sheet_name_task, startrow=start_row, startcol=0, index=False, header=False)
    df_pm[pm_plan_column].to_excel(writer, sheet_name=sheet_name_pm, startrow=start_row, startcol=0, index=False, header=False)
    df_workorder.to_excel(writer, sheet_name=sheet_name_wo, startrow=(start_row + start_offset), startcol=1, index=False, header=False)

def decorate_sheet_labor(sheet, thin_border, fill_color, yellow_fill):
    # ตกแต่งชีต JPPLAN-LABOR
    headers = ['GROUP', 'MOD', 'COMMENT']
    for i, header in enumerate(headers):
        cell = sheet.cell(row=2, column=11 + i, value=header)
        cell.border = thin_border

    for row in range(3, sheet.max_row + 1):
        for col in range(1, 11):
            cell = sheet.cell(row=row, column=col)
            cell.border = thin_border

            if sheet.cell(row=row, column=12).value == 0:
                sheet.cell(row=row, column=col).fill = fill_color

            if sheet.cell(row=row, column=13).value:
                sheet.cell(row=row, column=col).fill = yellow_fill

    sheet.insert_cols(idx=11)
    sheet.auto_filter.ref = "A2:N2"

# ฟังก์ชันสำหรับตกแต่งชีต JPPLAN-TASK
def decorate_sheet_task(sheet, thin_border, fill_color, yellow_fill, start_row, df_jop_plan_master):
    headers = ['START_DATE', 'FINISH_DATE', 'GROUP', 'MOD', 'COMMENT']
    for row in range(start_row + 1, start_row + 1 + len(df_jop_plan_master)):
        sheet[f'L{row}'].number_format = 'YYYY-MM-DD'
        sheet[f'M{row}'].number_format = 'YYYY-MM-DD'
    
    for i, header in enumerate(headers):
        cell = sheet.cell(row=2, column=12 + i, value=header)
        cell.border = thin_border

    for row in range(3, sheet.max_row + 1):
        for col in range(1, 12):
            cell = sheet.cell(row=row, column=col)
            cell.border = thin_border

            if sheet.cell(row=row, column=15).value == 0:
                sheet.cell(row=row, column=col).fill = fill_color

            if sheet.cell(row=row, column=16).value:
                sheet.cell(row=row, column=col).fill = yellow_fill

    sheet.insert_cols(idx=12)
    sheet.column_dimensions['M'].width = 20.0
    sheet.column_dimensions['N'].width = 20.0
    sheet.auto_filter.ref = "A2:Q2"

# ฟังก์ชันสำหรับตกแต่งชีต PM-PLAN
def decorate_sheet_pm(sheet, thin_border, fill_color, yellow_fill, start_row, df_pm_plan3_master):
    headers = ['MOD', 'MAIN_SYSTEM', 'MAIN_SYSTEM_DESC', 'UNIT_TYPE',
            'TYPE','SUB_SYSTEM', 'EQUIPMENT', 'SUB_SYSTEM_DESC',
            'KKS_NEW_DESC', 'GROUP', 'FINISH_DATE', 'FINISH TIME', 'COMMENT']
    
    for row in range(start_row + 1, start_row + 1 + len(df_pm_plan3_master)):
        sheet[f'O{row}'].number_format = 'YYYY-MM-DD'
        sheet[f'U{row}'].number_format = 'hh:mm:ss'
        sheet[f'AI{row}'].number_format = 'YYYY-MM-DD'
    
    for i, header in enumerate(headers):
        cell = sheet.cell(row = 2, column = 25 + i, value = header)
        cell.border = thin_border

    for row in range(3, sheet.max_row + 1):
        for col in range(1,25):
            cell = sheet.cell(row=row, column=col)
            cell.border = thin_border
            
            if sheet.cell(row=row, column=25).value == 0:
                sheet.cell(row=row, column=col).fill = fill_color

            if not sheet.cell(row=row, column=13).value:
                sheet.cell(row=row, column=col).font = Font(bold=True)
                
            if sheet.cell(row=row, column=37).value:
                sheet.cell(row=row, column=col).fill = yellow_fill

    sheet.insert_cols(idx=25)
    sheet.column_dimensions['AA'].width = 15.0
    # sheet.column_dimensions['AB'].width = 51.0
    sheet.column_dimensions['AE'].width = 15.0
    # sheet.column_dimensions['AG'].width = 51.0
    # sheet.column_dimensions['AH'].width = 90.0
    sheet.column_dimensions['AJ'].width = 15.0
    sheet.column_dimensions['AK'].width = 15.0

    sheet.auto_filter.ref = 'A2:AL2'

# ฟังก์ชันสำหรับตกแต่งชีต WO
# def decorate_sheet_wo(sheet):
#     # เพิ่มการตกแต่งสำหรับชีต WO ตามต้องการ
#     pass

def decorate_sheet(sheet, sheet_name, thin_border, fill_color, yellow_fill, start_row, df_jop_plan_master, df_pm_plan3_master):
    if "JPPLAN-LABOR" in sheet_name:
        decorate_sheet_labor(sheet, thin_border, fill_color, yellow_fill)
    elif "JPPLAN-TASK" in sheet_name:
        decorate_sheet_task(sheet, thin_border, fill_color, yellow_fill, start_row, df_jop_plan_master)  # ส่ง df_jop_plan_master มาด้วย
    elif "PMPlan" in sheet_name:
        decorate_sheet_pm(sheet, thin_border, fill_color, yellow_fill, start_row, df_pm_plan3_master)
    # elif "WO" in sheet_name:
    #     decorate_sheet_wo(sheet)