import os
import logging
# import shutil
import tkinter as tk
from tkinter import messagebox
from db.executor import DBExecutor
from excel.handler import ExcelHandler
from config.settings import (INPUT_FOLDER, RESULT_FOLDER,
                             ARCHIVE_FOLDER, WARNING_FOLDER,
                             RESULT_LOCAL_FOLDER,
                             USERNAME, ROWS_QUANTITY)


def check_folders(sftp):
    # if not os.path.exists(INPUT_FOLDER):
    #     os.makedirs(INPUT_FOLDER)
    if not os.path.exists(WARNING_FOLDER):
        os.makedirs(WARNING_FOLDER)
    if not os.path.exists(ARCHIVE_FOLDER):
        os.makedirs(ARCHIVE_FOLDER)
    if not os.path.exists(RESULT_LOCAL_FOLDER):
        os.makedirs(RESULT_LOCAL_FOLDER)

    # folders = [INPUT_FOLDER, WARNING_FOLDER, ARCHIVE_FOLDER, RESULT_FOLDER]
    folders = [INPUT_FOLDER, RESULT_FOLDER]
    for folder in folders:
        try:
            sftp.stat(folder)  # Check if the folder exists
        except FileNotFoundError:
            sftp.mkdir(folder)  # If it doesn't exist, create it


def create_log_file(sftp):
    # log_path = f'{RESULT_FOLDER}/{USERNAME}.log'
    # with sftp.file(log_path, 'w') as log_file:
    #     logging.basicConfig(stream=log_file,
    #                         level=logging.INFO,
    #                         format='%(asctime)s - %(levelname)s - %(message)s',
    #                         force=True)

    logging.basicConfig(filename=f'{RESULT_LOCAL_FOLDER}/{USERNAME}.log',
                        level=logging.INFO,
                        format='%(asctime)s - %(levelname)s - %(message)s',
                        force=True)


def show_alert(message):
    root = tk.Tk()  # Create new window
    root.withdraw()  # Hide main window
    # Show message
    messagebox.showinfo("File Processing", message)
    # Update and destroy window after showing message
    root.update()
    root.destroy()


# Remove a file if it already exists
def safe_move(sftp, old_path, new_path):
    try:
        try:
            sftp.stat(new_path)  # Check if the file exists
            sftp.remove(new_path)
        except FileNotFoundError:
            # If the file doesn't exist, do nothing
            pass
        sftp.rename(old_path, new_path)
    except Exception as e:
        print(f'Error moving file: {old_path} in {new_path}: {str(e)}')
        logging.info(f'Error moving file: {old_path} in {new_path}: {str(e)}')


def move_from_sftp_to_local(sftp, old_path, new_path):
    try:
        sftp.get(old_path, new_path)
        sftp.remove(old_path)
    except Exception as e:
        print(f'Error moving file: {old_path} in {new_path}: {str(e)}')
        logging.info(f'Error moving file: {old_path} in {new_path}: {str(e)}')


def process_files(sftp, files):
    all_warning_numbers = set()
    warning_name_files = ''
    errors = ''
    try:
        print(f'===============================================')
        logging.info(f'===============================================')
        print(files)
        logging.info(files)
        current_rows_quantity = ROWS_QUANTITY
        for k, f in enumerate(files):  # Do with each file
            remote_file_path = f'{INPUT_FOLDER}/{f}'
            # file_path = os.path.join(INPUT_FOLDER, f)
            print(f'-----------------------------------------------')
            logging.info(f'-----------------------------------------------')
            print(f'GET: {f}')
            logging.info(f'GET: {f}')
            try:
                with sftp.file(remote_file_path, 'rb') as remote_file:
                    # if os.path.isfile(file_path):
                    warning_numbers = ''
                    excel_handler = ExcelHandler(remote_file, remote_file_path)
                    ip_list = excel_handler.get_ip_list_from_xlsx_file()
                    errors += excel_handler.errors
                    excel_handler.errors = ''
                    excel_handler.create_output_xlsx_file(RESULT_FOLDER, RESULT_LOCAL_FOLDER)
                    errors += excel_handler.errors
                    excel_handler.errors = ''
                    db_executor = DBExecutor()
                    if db_executor.connect_on():
                        try:
                            DST_numbers_ls = []
                            for i, tuple_values in enumerate(ip_list):
                                DST_numbers = db_executor.execute(USERNAME, tuple_values)
                                errors += db_executor.errors
                                db_executor.errors = ''
                                print(f'response: {DST_numbers}')
                                logging.info(f'response: {DST_numbers}')
                                # if response is error then repeat request 5 time
                                if DST_numbers and next(iter(DST_numbers)) == 'ERROR':
                                    for j in range(5):
                                        DST_numbers = db_executor.execute(USERNAME, tuple_values)
                                        if DST_numbers and next(iter(DST_numbers)) != 'ERROR':
                                            break
                                    db_executor.errors = ''
                                DST_numbers_ls.append(DST_numbers)
                                if DST_numbers and next(iter(DST_numbers)) != 'ERROR':
                                    warning_numbers = db_executor.execute_check_numbers(DST_numbers)
                                    errors += db_executor.errors
                                    db_executor.errors = ''
                                    if warning_numbers:
                                        warning_name_files = (
                                            ' '.join([warning_name_files, f]))
                                        all_warning_numbers |= warning_numbers
                                        print(f'!!! WARNING NUMBERS: {warning_numbers} !!!')
                                        logging.info(f'!!! WARNING NUMBERS: {warning_numbers} !!!')
                                if i + 1 == current_rows_quantity or i + 1 == len(ip_list):
                                    excel_handler.save_result_to_output_xlsx_file(DST_numbers_ls)
                                    errors += excel_handler.errors
                                    excel_handler.errors = ''
                                    DST_numbers_ls = []
                                    current_rows_quantity += ROWS_QUANTITY
                        finally:
                            db_executor.connect_off()
                            errors += db_executor.errors
                    else:
                        errors += db_executor.errors
                    if warning_numbers:
                        # moving from sftp to local
                        warning_file_path = os.path.join(WARNING_FOLDER,
                                                         os.path.basename(remote_file_path))
                        move_from_sftp_to_local(sftp, remote_file_path, warning_file_path)
                        # sftp moving
                        # warning_file_path = os.path.join(WARNING_FOLDER,
                        #                                  os.path.basename(remote_file_path))
                        # safe_move(sftp, remote_file_path, warning_file_path)
                        # local moving
                        # warning_file_path = os.path.join(WARNING_FOLDER, file_path)
                        # shutil.move(file_path, warning_file_path)
                    else:
                        # we give the result file
                        with sftp.file(excel_handler.xlsx_remote_output_file, 'wb') as remote_output_file:
                            with open(excel_handler.xlsx_output_file, 'rb') as local_file:
                                remote_output_file.write(local_file.read())
                        # moving from sftp to local
                        archive_file_path = os.path.join(ARCHIVE_FOLDER,
                                                         os.path.basename(remote_file_path))
                        move_from_sftp_to_local(sftp, remote_file_path, archive_file_path)
                        # sftp moving
                        # archive_file_path = os.path.join(ARCHIVE_FOLDER,
                        #                                  os.path.basename(remote_file_path))
                        # safe_move(sftp, remote_file_path, archive_file_path)
                        # local moving
                        # archive_file_path = os.path.join(ARCHIVE_FOLDER, os.path.basename(file_path))
                        # shutil.move(file_path, archive_file_path)
            except FileNotFoundError:
                print(f"File not found: {remote_file_path}")
                logging.error(f"File not found: {remote_file_path}")
                continue
    except Exception as e:
        print(f'General error:\n {str(e)}')
        logging.error(f'General error:\n {str(e)}')
        errors += 'general error\n'
    finally:
        if all_warning_numbers:
            formatted_numbers = "\n".join(all_warning_numbers)
            show_alert(f"WARNING NUMBERS in files {warning_name_files}:\n"
                       f"{formatted_numbers}")
        if errors:
            show_alert(f"ERRORS:\n"
                       f"{errors}")
