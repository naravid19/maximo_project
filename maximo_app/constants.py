"""
Constants for Maximo App.

This module contains all hardcoded values used throughout the application,
centralized for easy maintenance and configuration.
"""

# ==============================================================================
# Organization and Status Defaults
# ==============================================================================
ORGID = 'EGAT'
PLUSC_REVNUM = 0
DEFAULT_STATUS = 'ACTIVE'
PLUSC_JP_REVNUM = 0
FREQ_UNIT = 'YEARS'
LEAD_TIME = 7


# ==============================================================================
# Column Length Limits (Based on Maximo Database Schema)
# ==============================================================================
MAX_KKS_LENGTH = 30
MAX_EQUIPMENT_LENGTH = 100
MAX_TASK_LENGTH = 100
MAX_ROUTE_LENGTH = 20
MAX_CRAFT_LENGTH = 12
MAX_RESPONSE_LENGTH = 12
MAX_PTW_LENGTH = 250
MAX_DURATION_DIGITS = 8
MAX_TASK_ORDER_LENGTH = 12
MAX_JPNUM_LENGTH = 30
MAX_PMNUM_LENGTH = 30


# ==============================================================================
# Required Columns for Schedule File
# ==============================================================================
SCHEDULE_REQUIRED_COLUMNS = [
    'KKS', 'EQUIPMENT', 'TASK_XX', 'TASK',
    'RESPONSE', 'ROUTE', 'DURATION_(HR.)', 'START_DATE', 'FINISH_DATE',
    'SUPERVISOR', 'FOREMAN', 'SKILL', 'RESPONSE_CRAFT',
    'ประเภทของ_PERMIT_TO_WORK', 'TYPE', 'COMMENT'
]

SCHEDULE_IMPORTANT_COLUMNS = ['TASK_XX']

SCHEDULE_USE_COLUMNS = [
    'KKS', 'EQUIPMENT', 'TASK_XX', 'TASK',
    'RESPONSE', 'ROUTE', 'DURATION_(HR.)', 'START_DATE', 'FINISH_DATE',
    'SUPERVISOR', 'FOREMAN', 'SKILL', 'RESPONSE_CRAFT',
    'ประเภทของ_PERMIT_TO_WORK', 'TYPE'
]

LOCATION_COLUMNS = ['Location', 'Description']


# ==============================================================================
# Valid Values
# ==============================================================================
VALID_TYPES = ['ME', 'EE', 'CV', 'IC']


# ==============================================================================
# Excel Styling Colors (RGB Hex values)
# ==============================================================================
COLOR_RED = "FF0000"
COLOR_YELLOW = "FFFF00"
COLOR_BLUE = "00B0F0"


# ==============================================================================
# Error Messages (Thai)
# ==============================================================================
ERROR_MESSAGES = {
    'equipment_too_long': 'EQUIPMENT มีความยาวมากกว่า 100 ตัวอักษร',
    'task_too_long': 'TASK มีความยาวมากกว่า 100 ตัวอักษร',
    'no_task': 'ไม่มี TASK',
    'route_too_long': 'ROUTE มีความยาวมากกว่า 20 ตัวอักษร',
    'invalid_duration': 'DURATION_(HR.) ไม่ถูกต้อง',
    'start_date_has_letters': 'START_DATE มีตัวอักษร',
    'finish_date_has_letters': 'FINISH_DATE มีตัวอักษร',
    'invalid_start_date': 'START_DATE ไม่ถูกต้อง',
    'invalid_finish_date': 'FINISH_DATE ไม่ถูกต้อง',
    'kks_too_long': 'KKS มีความยาวมากกว่า 30 ตัวอักษร',
    'plant_unit_mismatch': 'Plant Unit ไม่สอดคล้อง',
    'no_start_date': 'ไม่มี START_DATE',
    'no_finish_date': 'ไม่มี FINISH_DATE',
    'no_skill_rate': 'ไม่มี SKILL RATE (จำเป็นต้องกรอก)',
    'invalid_skill_rate': 'SKILL RATE ไม่ถูกต้อง',
    'no_response_craft': 'ไม่มี RESPONSE_CRAFT',
    'craft_too_long': 'RESPONSE_CRAFT มีความยาวมากกว่า 12 ตัวอักษร',
    'no_response': 'ไม่มี RESPONSE',
    'response_too_long': 'RESPONSE มีความยาวมากกว่า 12 ตัวอักษร',
    'kks_not_found': 'ไม่พบ kks',
    'no_duration': 'ไม่มี DURATION_(HR.)',
    'duration_too_long': 'DURATION_(HR.) มีความยาวมากกว่า 8 หลัก',
    'no_ptw': 'ไม่มี ประเภทของ_PERMIT_TO_WORK',
    'ptw_too_long': 'ประเภทของ_PERMIT_TO_WORK มีความยาวมากกว่า 250 ตัวอักษร',
    'no_type': 'ไม่มี TYPE (จำเป็นต้องกรอก)',
    'invalid_type': 'TYPE ไม่ถูกต้อง',
    'invalid_task_order': 'TASK_ORDER ไม่ถูกต้อง',
    'no_task_order': 'ไม่มี TASK_ORDER',
    'task_order_too_long': 'TASK_ORDER มีความยาวมากกว่า 12 ตัวอักษร',
    'no_kks_required': 'ไม่มี KKS (จำเป็นต้องกรอก)',
    'no_equipment_required': 'ไม่มี EQUIPMENT (จำเป็นต้องกรอก)',
    'jpnum_too_long': 'JPNUM มีความยาวมากกว่า 30 ตัวอักษร',
    'description_too_long': 'DESCRIPTION มีความยาวมากกว่า 100 ตัวอักษร',
    'jobtask_too_long': 'JOBTASK มีความยาวมากกว่า 100 ตัวอักษร',
    'pmnum_too_long': 'PMNUM มีความยาวมากกว่า 30 ตัวอักษร',
}
