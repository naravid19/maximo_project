# Path : maximo_project\maximo_app\tasks.py

from background_task import background
import os
import time
@background(schedule=900)  # 900 วินาที = 15 นาที
# พารามิเตอร์ที่กำหนดให้ Task นี้ถูกรันหลังจาก 900 วินาที (หรือ 15 นาที) หลังจากที่มันถูกเรียกใช้ครั้งแรก

def delete_old_files_task():
    now = time.time()
    temp_dir = 'temp'
    age_limit = 15 * 60

    for filename in os.listdir(temp_dir):
        file_path = os.path.join(temp_dir, filename)
        if os.path.isfile(file_path):
            file_age = now - os.path.getmtime(file_path)
            if file_age > age_limit:
                os.remove(file_path)
