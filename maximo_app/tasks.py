# path : maximo_project\maximo_app\tasks.py

import os
import time
import logging
from background_task import background
from django.conf import settings

# ตั้งค่า Logging
LOG_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"
logging.basicConfig(level=logging.INFO, format=LOG_FORMAT)

# ดึงค่าจาก settings
TEMP_DIR = settings.TEMP_DIR
AGE_LIMIT = settings.FILE_AGE_LIMIT

# กำหนด path ของ lock file เพื่อป้องกันการรัน task ซ้ำ
LOCK_FILE = os.path.join(TEMP_DIR, ".cleanup.lock")

@background(schedule=3600)  # รันทุก 1 ชั่วโมง
def delete_old_files_task():
    now = time.time()
    if not os.path.exists(TEMP_DIR):
        logging.warning(f"Folder '{TEMP_DIR}' does not exist. Skipping cleanup.")
        return

    # ป้องกันการทำงานซ้ำด้วย lock file
    if os.path.exists(LOCK_FILE):
        logging.info("Another cleanup task is already running. Skipping this run.")
        return

    # สร้าง lock file
    try:
        with open(LOCK_FILE, "w") as lock_file:
            lock_file.write(str(now))
    except Exception as e:
        logging.error(f"Error creating lock file: {e}")
        return

    try:
        for filename in os.listdir(TEMP_DIR):
            file_path = os.path.join(TEMP_DIR, filename)
            
            if os.path.isfile(file_path):
                try:
                    file_age = now - os.path.getmtime(file_path)
                except Exception as e:
                    logging.error(f"Error getting modification time for {file_path}: {e}")
                    continue

                if file_age > AGE_LIMIT:
                    try:
                        os.remove(file_path)
                        logging.info(f"Deleted old file: {file_path}")
                    except PermissionError:
                        logging.error(f"Permission denied while deleting {file_path}")
                    except FileNotFoundError:
                        logging.error(f"File not found: {file_path}")
                    except Exception as e:
                        logging.error(f"Error deleting file {file_path}: {e}")
    except Exception as e:
        logging.error(f"Error during cleanup process: {e}")
    finally:
        try:
            if os.path.exists(LOCK_FILE):
                os.remove(LOCK_FILE)
        except Exception as e:
            logging.error(f"Error removing lock file: {e}")
