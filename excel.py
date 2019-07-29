#!/usr/bin/python3

import sys
import openpyxl

# DECLARE CONSTANTS
XLS = ".xls"
XLSX = ".xlsx"
GOOD = "_good_"
WRONG = "_wrong_"
THRESHOLD = "threshold-"
TXT = ".txt"
PERCENTAGE_THRESHOLD = 25
MIN_ROW = 0
MAX_ROW = 27
MIN_COL = 2
# ~~ MAX_COL is set automatically below after opening the excel file


# DECLARE VARIABLES
columns_index = []
columns_info_wrong = {}
columns_info_good = {}


# DECLARE FUNCTIONS
def check_input_file_extension(file):
    if not file.lower().endswith((XLS, XLSX)):
        sys.exit('File extension is missing or file is not an excel file')


def create_output_excel_file_name(file):
    filename = file.split(".")[0]
    file_extension = file.split(".")[1]
    return filename + "_" + THRESHOLD + str(PERCENTAGE_THRESHOLD) + "." + file_extension


def create_report_file_name(file, good_or_wrong):
    filename = file.split(".")[0]
    return filename + good_or_wrong + THRESHOLD + str(PERCENTAGE_THRESHOLD) + TXT


def calculate_percentage_difference(first_value, second_value):
    try:
        result = (abs(first_value - second_value) / second_value) * 100.0
    except ZeroDivisionError:
        result = 0
    return result


def delete_columns(sheet, columns_list):
    index = 0
    for column in columns_list:
        sheet.delete_cols(column - index)
        index = index + 1


def write_columns_info_to_file(filename, columns_info):
    with open(filename, 'w') as f:
        for k, v in columns_info.items():
            f.write(k + ":" + " " + str(v) + "\n")


def main(excel_file):
    # 1- CHECK INPUT FILE IS CORRECT AND PREPARE NAME OF OUTPUT FILES
    check_input_file_extension(excel_file)
    excel_output_file = create_output_excel_file_name(excel_file)
    report_file_wrong = create_report_file_name(excel_file, WRONG)
    report_file_good = create_report_file_name(excel_file, GOOD)

    # 2- OPEN THE EXCEL FILE
    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook.active
    MAX_COL = sheet.max_column

    # 3- ITERATE THROUGH COLUMNS AND IDENTIFY BAD COLUMNS GIVEN THE THRESHOLD PERCENTAGE
    for col in sheet.iter_cols(min_row=MIN_ROW, min_col=MIN_COL, max_row=MAX_ROW, max_col=MAX_COL):
        first_cell_value = col[MIN_ROW + 1].value
        second_cell_value = col[MAX_ROW - 1].value
        difference = calculate_percentage_difference(first_cell_value, second_cell_value)
        if difference > PERCENTAGE_THRESHOLD:
            columns_index.append(col[0].column)
            columns_info_wrong.update({col[0].value: difference})
        else:
            columns_info_good.update({col[0].value: difference})

    # 4- DELETE THE BAD COLUMNS
    if len(columns_index) != 0:
        delete_columns(sheet, columns_index)
        write_columns_info_to_file(report_file_good, columns_info_good)
        write_columns_info_to_file(report_file_wrong, columns_info_wrong)
        workbook.save(excel_output_file)
        print(excel_file + " DONE! " + str(len(columns_index)) + " columns removed")
    else:
        print(excel_file + " DONE! " + " No column to delete")


#########################
## PROGRAM STARTS HERE ##
#########################

# CHECK IF PERCENTAGE IS GIVEN AS ARGUMENT
if len(sys.argv) > 2:
    PERCENTAGE_THRESHOLD = int(sys.argv[2])

# CHECK IF PROGRAM SHOULD PROCESS A EXCEL FILE OR A LIST OF EXCEL FILES
if sys.argv[1].lower().endswith(TXT):
    with open(sys.argv[1]) as fp:
        line = fp.readline().rstrip('\n')
        while line:
            main(line)
            # GO NEXT LINE
            line = fp.readline().rstrip('\n')
            # REINITIALIZE VARIABLES
            columns_index = []
            columns_info_wrong = {}
            columns_info_good = {}
else:
    main(sys.argv[1])
