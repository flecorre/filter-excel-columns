#!/usr/bin/python3

import sys
import datetime
import openpyxl

# DECLARE CONSTANTS
XLS = ".xls"
XLSX = ".xlsx"
GOOD = "_GOOD_"
WRONG = "_WRONG_"
THRESHOLD = "THRESHOLD-"
TODAY = str(datetime.date.today())
TXT = ".txt"
PERCENTAGE_THRESHOLD = 25
MIN_ROW = 0
MAX_ROW = 27
MIN_COL = 2
# ~~ MAX_COL is set automatically below after opening the excel file


# DECLARE VARIABLES
columns_index = []  # a list
columns_info_wrong = {}  # a dictionary
columns_info_good = {}


# DECLARE FUNCTIONS
def check_input_file_extension(file):
    if not file.lower().endswith((XLS, XLSX)):
        sys.exit('File extension is missing or file is not an excel file')


def create_output_excel_file_name(file):
    filename = file.split(".")[0]
    file_extension = file.split(".")[1]
    return filename + "_" + THRESHOLD + str(PERCENTAGE_THRESHOLD) + "_" + TODAY + "." + file_extension


def create_report_file_name(file, good_or_wrong):
    filename = file.split(".")[0]
    return filename + good_or_wrong + THRESHOLD + str(PERCENTAGE_THRESHOLD) + "_" + TODAY + TXT


def calculate_percentage_difference(first_value, second_value):
    try:
        result = (abs(first_value - second_value) / second_value) * 100.0
    except ZeroDivisionError:
        result = 0
    return result


def delete_columns(columns_list):
    index = 0
    for column in columns_list:
        sheet.delete_cols(column - index)
        index = index + 1


def write_columns_info_to_file(filename, columns_info):
    with open(filename, 'w') as f:
        for k, v in columns_info.items():
            f.write(k+":" + " " + str(v) + "\n")


# CODE BEGINS HERE

# 1- CHECK IF PERCENTAGE IS GIVEN AS ARGUMENT
if len(sys.argv) > 2:
    PERCENTAGE_THRESHOLD = int(sys.argv[2])


# 2- CHECK INPUT FILE IS CORRECT AND PREPARE NAME OF OUTPUT FILES
excel_input_file = sys.argv[1]  # take first python argument as an excel file
check_input_file_extension(excel_input_file)
excel_output_file = create_output_excel_file_name(excel_input_file)
report_file_wrong = create_report_file_name(excel_input_file, WRONG)
report_file_good = create_report_file_name(excel_input_file, GOOD)


# 3- OPEN EXCEL FILE
workbook = openpyxl.load_workbook(excel_input_file)
sheet = workbook.active
MAX_COL = sheet.max_column


# 4- ITERATE THROUGH COLUMNS AND IDENTIFY WRONG COLUMNS
for col in sheet.iter_cols(min_row=MIN_ROW, min_col=MIN_COL, max_row=MAX_ROW, max_col=MAX_COL):
    first_cell_value = col[MIN_ROW + 1].value
    second_cell_value = col[MAX_ROW - 1].value
    difference = calculate_percentage_difference(first_cell_value, second_cell_value)
    if difference > PERCENTAGE_THRESHOLD:
        columns_index.append(col[0].column)
        columns_info_wrong.update({col[0].value: difference})
    else:
        columns_info_good.update({col[0].value: difference})


# 5- DELETE WRONG COLUMNS
if len(columns_index) != 0:
    delete_columns(columns_index)
    write_columns_info_to_file(report_file_good, columns_info_good)
    write_columns_info_to_file(report_file_wrong, columns_info_wrong)
    workbook.save(excel_output_file)
    print("DONE! " + str(len(columns_index)) + " columns removed")
else:
    sys.exit("No column to delete")
