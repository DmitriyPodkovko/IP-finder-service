import os
import time
from config.settings import INTERVAL, INPUT_FOLDER
from processor.handler import (check_folders,
                               create_log_file,
                               process_files)


# default interval is 2 min (120 sec) if not defined in setting.py
def main(interval=120):
    # Define the environment variable DJANGO_SETTINGS_MODULE
    # for DATABASES settings configuration
    os.environ.setdefault('DJANGO_SETTINGS_MODULE',
                          'config.settings')
    while True:
        check_folders()
        create_log_file()
        files = os.listdir(INPUT_FOLDER)
        if files:
            process_files(files)
        time.sleep(interval)


if __name__ == "__main__":
    main(INTERVAL)
