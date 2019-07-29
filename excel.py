#!/usr/bin/python3

import sys
import openpyxl
import logging

### CONSTANTS ###
STARS = "*******************************************************"
XLS = ".xls"
XLSX = ".xlsx"
GOOD = "_good_"
WRONG = "_wrong_"
FILTERED = "filtered_"
THRESHOLD = "threshold-"
BACKGROUND_SUBTRACTED = "background-subtracted"
TXT = ".txt"
PERCENTAGE_THRESHOLD = 25
BACKGROUND_COLUMN_INDEX = 2
MIN_ROW = 0
MAX_ROW = 27
MIN_COL = 2
# ~~ MAX_COL is set automatically below after opening the excel file


### VARIABLES ###
columns_index = []
columns_info_wrong = {}
columns_info_good = {}
excel_background_subtracted_file = ""
excel_filtered_file = ""
report_file_wrong = ""
report_file_good = ""
workbook = None
sheet = None


# SET LOGGER
logging.basicConfig(format='%(asctime)s - %(message)s', level=logging.INFO)


### FUNCTIONS ###
def prepare_output_files(excel_file):
    global excel_background_subtracted_file
    global excel_filtered_file
    global report_file_wrong
    global report_file_good
    check_input_file_extension(excel_file)
    excel_background_subtracted_file = create_output_excel_file_name(excel_file, BACKGROUND_SUBTRACTED)
    filtered_threshold = FILTERED + THRESHOLD + str(PERCENTAGE_THRESHOLD)
    excel_filtered_file = create_output_excel_file_name(excel_file, filtered_threshold)
    report_file_wrong = create_report_file_name(excel_file, WRONG)
    report_file_good = create_report_file_name(excel_file, GOOD)


def open_excel_file(excel_file):
    global workbook
    global sheet
    logging.info("********* {} *********".format(excel_file))
    logging.info("opening file...")
    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook.active


def check_input_file_extension(file):
    if not file.lower().endswith((XLS, XLSX)):
        sys.exit('File extension is missing or file is not an excel file')


def create_output_excel_file_name(file, suffix):
    filename = file.split(".")[0]
    file_extension = file.split(".")[1]
    return filename + "_" + suffix + "." + file_extension


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
    logging.info("{} columns removed...".format(str(len(columns_list))))


def write_columns_info_to_file(filename, columns_info):
    with open(filename, 'w') as f:
        for k, v in columns_info.items():
            f.write(k + ":" + " " + str(v) + "\n")


def reinit_variables():
    global columns_index
    global columns_info_wrong
    global columns_info_good
    global excel_background_subtracted_file
    global excel_filtered_file
    global report_file_wrong
    global report_file_good
    global workbook
    global sheet
    columns_index = []
    columns_info_wrong = {}
    columns_info_good = {}
    excel_background_subtracted_file = ""
    excel_filtered_file = ""
    report_file_wrong = ""
    report_file_good = ""
    workbook = None
    sheet = None


def subtract_background(workbook, sheet):
    logging.info("subtracting background...")
    for row in sheet.iter_rows(min_row=2, min_col=BACKGROUND_COLUMN_INDEX):
        background_cell = row[0].value
        for cell in row:
            cell.value = cell.value - background_cell
    sheet.delete_cols(BACKGROUND_COLUMN_INDEX)
    logging.info("writing background filtered file: '{}'".format(excel_background_subtracted_file))
    workbook.save(excel_background_subtracted_file)


def filter_columns(workbook, sheet):
    MAX_COL = sheet.max_column

    # ITERATE THROUGH COLUMNS AND IDENTIFY BAD COLUMNS GIVEN THE THRESHOLD PERCENTAGE
    logging.info("calculating threshold for every ROI...")
    for col in sheet.iter_cols(min_row=MIN_ROW, min_col=MIN_COL, max_row=MAX_ROW, max_col=MAX_COL):
        first_cell_value = col[MIN_ROW + 1].value
        second_cell_value = col[MAX_ROW - 1].value
        difference = calculate_percentage_difference(first_cell_value, second_cell_value)
        if difference > PERCENTAGE_THRESHOLD:
            columns_index.append(col[0].column)
            columns_info_wrong.update({col[0].value: difference})
        else:
            columns_info_good.update({col[0].value: difference})

    # DELETE COLUMNS
    logging.info("deleting columns...")
    if len(columns_index) != 0:
        delete_columns(sheet, columns_index)
        logging.info("writing report files: '{}' and '{}'".format(report_file_good, report_file_wrong))
        write_columns_info_to_file(report_file_good, columns_info_good)
        write_columns_info_to_file(report_file_wrong, columns_info_wrong)
        logging.info("writing filtered file: '{}'".format(excel_filtered_file))
        workbook.save(excel_filtered_file)
    else:
        logging.critical("no column to delete...")


def main(line):
    prepare_output_files(line)
    open_excel_file(line)
    subtract_background(workbook, sheet)
    filter_columns(workbook, sheet)
    reinit_variables()
    logging.info(STARS)


#########################
## PROGRAM STARTS HERE ##
#########################

# CHECK IF PERCENTAGE IS GIVEN AS ARGUMENT
if len(sys.argv) > 2:
    PERCENTAGE_THRESHOLD = int(sys.argv[2])
logging.info("************** THRESHOLD IS SET TO: {} % **************".format(str(PERCENTAGE_THRESHOLD)))

# CHECK IF PROGRAM SHOULD PROCESS A EXCEL FILE OR A LIST OF EXCEL FILES
if sys.argv[1].lower().endswith(TXT):
    with open(sys.argv[1]) as fp:
        line = fp.readline().rstrip('\n')
        while line:
            main(line)
            line = fp.readline().rstrip('\n')
else:
    main(sys.argv[1])
