#!/usr/bin/python3

import openpyxl
import logging
import argparse


### CONSTANTS ###
STARS = "******************************************************"
XLS = ".xls"
XLSX = ".xlsx"
GOOD = "_good_"
WRONG = "_wrong_"
FILTERED = "filtered_"
THRESHOLD = "threshold-"
BACKGROUND_SUBTRACTED = "bg-subtracted"
TXT = ".txt"
BACKGROUND_MIN_ROW = 2
BACKGROUND_COLUMN_INDEX = 2
FILTER_MIN_ROW = 0
FILTER_MAX_ROW = 27
FILTER_MIN_COL = 2
# ~~ FILTER_MAX_COL is set automatically below before filtering


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
percentage_threshold = None
skip_background = None


# SET LOGGER
logging.basicConfig(format='%(asctime)s - %(message)s', level=logging.INFO)


### FUNCTIONS ###

def valid_arg_list(param):
    if not(param.lower().endswith(TXT)):
        msg = "{} is not a valid list. Should be a .txt file".format(param)
        raise argparse.ArgumentTypeError(msg)
    return param


def valid_arg_excel(param):
    if not(param.lower().endswith(XLS) or param.lower().endswith(XLSX)):
        msg = "{} is not a excel file. Should be a .xls or .xlsx file".format(param)
        raise argparse.ArgumentTypeError(msg)
    return param


def valid_arg_threshold(param):
    try:
        if not(0 <= int(param) <= 100):
            raise ValueError
    except ValueError:
        msg = "Threshold value should be a number between 0 and 100"
        raise argparse.ArgumentTypeError(msg)
    return param


def prepare_output_files(excel_file):
    global excel_background_subtracted_file
    global excel_filtered_file
    global report_file_wrong
    global report_file_good
    check_file_extension(excel_file)
    excel_background_subtracted_file = create_output_excel_file_name(excel_file, BACKGROUND_SUBTRACTED)
    filtered_threshold = FILTERED + THRESHOLD + str(percentage_threshold)
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


def check_file_extension(file):
    if not file.lower().endswith((XLS, XLSX)):
        raise TypeError('File extension is missing or file is not an excel file')


def create_output_excel_file_name(file, suffix):
    filename = file.split(".")[0]
    file_extension = file.split(".")[1]
    return filename + "_" + suffix + "." + file_extension


def create_report_file_name(file, good_or_wrong):
    filename = file.split(".")[0]
    return filename + good_or_wrong + THRESHOLD + str(percentage_threshold) + TXT


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


def reinit_excel_variables():
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
    for row in sheet.iter_rows(min_row=BACKGROUND_MIN_ROW, min_col=BACKGROUND_COLUMN_INDEX):
        background_cell = row[0].value
        for cell in row:
            cell.value = cell.value - background_cell
    sheet.delete_cols(BACKGROUND_COLUMN_INDEX)
    logging.info("writing background filtered file: '{}'".format(excel_background_subtracted_file))
    workbook.save(excel_background_subtracted_file)


def filter_columns(workbook, sheet):
    FILTER_MAX_COL = sheet.max_column

    # ITERATE THROUGH COLUMNS AND IDENTIFY BAD COLUMNS GIVEN THE THRESHOLD PERCENTAGE
    logging.info("calculating threshold for every ROI...")
    for col in sheet.iter_cols(min_row=FILTER_MIN_ROW, min_col=FILTER_MIN_COL, max_row=FILTER_MAX_ROW, max_col=FILTER_MAX_COL):
        first_cell_value = col[FILTER_MIN_ROW + 1].value
        second_cell_value = col[FILTER_MAX_ROW - 1].value
        difference = calculate_percentage_difference(first_cell_value, second_cell_value)
        if difference > percentage_threshold:
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
    if not skip_background:
        subtract_background(workbook, sheet)
    filter_columns(workbook, sheet)
    reinit_excel_variables()
    logging.info(STARS)


#########################
## PROGRAM STARTS HERE ##
#########################

# PARSE ARGUMENTS
parser = argparse.ArgumentParser()
parser.add_argument('-sbg', '--skip-bg', dest='skip_background', action='store_true', default=False, help="skip background subtraction step")
parser.add_argument('-t', '--threshold', dest='threshold', type=valid_arg_threshold, default=25, help="override threshold value")
mutually_exclusive = parser.add_mutually_exclusive_group(required=True)
mutually_exclusive.add_argument('-e', '--excel', dest='excel_file', type=valid_arg_excel, help='process only one excel file')
mutually_exclusive.add_argument('-l', '--list', dest='excel_list', type=valid_arg_list, help='process a list of excel files declared in a .txt file,'
                                                                                             'only one file should be declared per line')
args = parser.parse_args()

# SET PERCENTAGE THRESHOLD VALUE (default is 25)
percentage_threshold = args.threshold
logging.info("************** THRESHOLD IS SET TO: {}% **************".format(str(percentage_threshold)))

skip_background = args.skip_background
if skip_background:
    logging.info("************** BACKGROUND SUBTRACTION STEP WILL BE SKIPPED **************")
logging.info(STARS)

# CHECK IF PROGRAM SHOULD PROCESS A EXCEL FILE OR A LIST OF EXCEL FILES
if args.excel_list is not None:
    with open(args.excel_list) as fp:
        line = fp.readline().rstrip('\n')
        while line:
            main(line)
            line = fp.readline().rstrip('\n')
elif args.excel_file is not None:
    main(args.excel_file)
