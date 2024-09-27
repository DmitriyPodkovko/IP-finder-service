import os
import logging
import shutil
import tkinter as tk
from tkinter import messagebox
from db.executor import DBExecutor
from excel.handler import ExcelHandler
from config.settings import (INPUT_FOLDER, RESULT_FOLDER,
                             ARCHIVE_FOLDER, WARNING_FOLDER,
                             USERNAME, ROWS_QUANTITY)


def check_folders():
    if not os.path.exists(INPUT_FOLDER):
        os.makedirs(INPUT_FOLDER)
    if not os.path.exists(WARNING_FOLDER):
        os.makedirs(WARNING_FOLDER)
    if not os.path.exists(ARCHIVE_FOLDER):
        os.makedirs(ARCHIVE_FOLDER)
    if not os.path.exists(RESULT_FOLDER):
        os.makedirs(RESULT_FOLDER)


def create_log_file():
    logging.basicConfig(filename=f'{RESULT_FOLDER}/{USERNAME}.log',
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


def process_files(files):
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
            file_path = os.path.join(INPUT_FOLDER, f)
            print(f'-----------------------------------------------')
            logging.info(f'-----------------------------------------------')
            print(f'GET: {f}')
            logging.info(f'GET: {f}')
            if os.path.isfile(file_path):
                warning_numbers = ''
                excel_handler = ExcelHandler(file_path)
                ip_list = excel_handler.get_ip_list_from_xlsx_file()
                errors += excel_handler.errors
                excel_handler.errors = ''
                excel_handler.create_output_xlsx_file(RESULT_FOLDER)
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
                    warning_file_path = os.path.join(WARNING_FOLDER, os.path.basename(file_path))
                    shutil.move(file_path, warning_file_path)
                else:
                    archive_file_path = os.path.join(ARCHIVE_FOLDER, os.path.basename(file_path))
                    shutil.move(file_path, archive_file_path)
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
