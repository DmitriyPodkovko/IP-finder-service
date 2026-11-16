import os
import time
import paramiko
from config.settings import (INTERVAL, INPUT_FOLDER,
                             SFTP_HOST, SFTP_PORT,
                             SFTP_USERNAME, SFTP_PASSWORD)
from processor.handler import (check_folders,
                               create_log_file,
                               process_files)


def connect_sftp():
    transport = paramiko.Transport((SFTP_HOST, SFTP_PORT))
    transport.connect(username=SFTP_USERNAME, password=SFTP_PASSWORD)
    sftp = paramiko.SFTPClient.from_transport(transport)
    return sftp, transport


# default interval is 2 min (120 sec) if not defined in setting.py
def main(interval=120):
    # Define the environment variable DJANGO_SETTINGS_MODULE
    # for DATABASES settings configuration
    os.environ.setdefault('DJANGO_SETTINGS_MODULE',
                          'config.settings')

    sftp, transport = connect_sftp()

    try:
        while True:
            check_folders(sftp)
            create_log_file(sftp)
            files = sftp.listdir(INPUT_FOLDER)
            if files:
                process_files(sftp, files)
            time.sleep(interval)
    finally:
        sftp.close()
        transport.close()


if __name__ == "__main__":
    main(INTERVAL)
