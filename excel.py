#!/usr/bin/python3

import openpyxl
import logging
import argparse


### CONSTANTS ###
STARS = "******************************************************"
XLS = ".xls"
XLSX = ".xlsx"
THRESHOLD = "threshold-"
TXT = ".txt"
SHEET_BACKGROUND_SUBTRACTED = "bg_subtracted"
SHEET_GOOD_ROI = "good ROI"
SHEET_WRONG_ROI = "wrong ROI"
MEAN_GOOD_ROI = "mean good ROI"
MEAN_WRONG_ROI = "mean wrong ROI"
RESULT_ABOVE = "result above"
RESULT_BELOW = "result below"
BACKGROUND_MIN_ROW = 2
BACKGROUND_COLUMN_INDEX = 2
FILTER_MIN_ROW = 0
FILTER_MAX_ROW = 21
FILTER_MIN_COL = 2
# ~~ FILTER_MAX_COL is set automatically below before filtering


### VARIABLES ###
# columns used during filtering
excel_output_file = ""
report_file_wrong = ""
report_file_good = ""
workbook = None
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
    global excel_output_file
    global report_file_wrong
    global report_file_good
    check_file_extension(excel_file)
    filtered_threshold = THRESHOLD + str(percentage_threshold)
    excel_output_file = create_output_excel_file_name(excel_file, filtered_threshold)


def check_file_extension(file):
    if not file.lower().endswith((XLS, XLSX)):
        raise TypeError('{} extension is missing or file is not an excel file'.format(file))


def create_output_excel_file_name(file, suffix):
    filename = file.split(".")[0]
    file_extension = file.split(".")[1]
    return filename + "_" + suffix + "." + file_extension


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


def write_data(sheet, columns_dict):
    next_row = 1
    for k, v in columns_dict.items():
        sheet.cell(column=1, row=next_row, value=k)
        sheet.cell(column=2, row=next_row, value=v)
        next_row += 1


def open_excel_file(excel_file):
    global workbook
    logging.info("opening '{}'...".format(excel_file))
    workbook = openpyxl.load_workbook(excel_file)


def copy_original_data(wb):
    sheet = wb.active
    copy_sheet = wb.copy_worksheet(sheet)
    copy_sheet.title = "original data"


def copy_worksheet(wb, sheet, title):
    new_worksheet = wb.copy_worksheet(sheet)
    new_worksheet.title = title
    return new_worksheet


def subtract_background(wb):
    sheet = wb.active
    subtracted_bg_sheet = copy_worksheet(wb, sheet, SHEET_BACKGROUND_SUBTRACTED)
    logging.info("subtracting background...")
    for row in subtracted_bg_sheet.iter_rows(min_row=BACKGROUND_MIN_ROW, min_col=BACKGROUND_COLUMN_INDEX):
        background_cell = row[0].value
        for cell in row:
            cell.value = cell.value - background_cell
    subtracted_bg_sheet.delete_cols(BACKGROUND_COLUMN_INDEX)


def filter_columns(wb):
    if not skip_background:
        sheet = wb[SHEET_BACKGROUND_SUBTRACTED]
    else:
        sheet = wb.active

    FILTER_MAX_COL = sheet.max_column

    # ITERATE THROUGH COLUMNS AND IDENTIFY BAD COLUMNS GIVEN THE THRESHOLD PERCENTAGE
    logging.info("filtering ROIs...")
    columns_index_wrong = []
    columns_index_good = []
    columns_info_wrong = {}
    columns_info_good = {}
    for col in sheet.iter_cols(min_row=FILTER_MIN_ROW, min_col=FILTER_MIN_COL, max_row=FILTER_MAX_ROW, max_col=FILTER_MAX_COL):
        first_cell_value = col[FILTER_MIN_ROW + 1].value
        second_cell_value = col[FILTER_MAX_ROW - 1].value
        difference = calculate_percentage_difference(first_cell_value, second_cell_value)
        if difference > percentage_threshold or difference < 0:
            columns_index_wrong.append(col[0].column)
            columns_info_wrong.update({col[0].value: difference})
        else:
            columns_index_good.append(col[0].column)
            columns_info_good.update({col[0].value: difference})

    # DELETE COLUMNS IF WRONG COLUMNS ARE FOUND
    if len(columns_index_wrong) != 0:
        logging.info("{} good columns found...".format(str(len(columns_index_good))))
        logging.info("{} wrong columns found...".format(str(len(columns_index_wrong))))
        logging.info("deleting columns...")
        sheet_good_roi = copy_worksheet(wb, sheet, SHEET_GOOD_ROI)
        sheet_wrong_roi = copy_worksheet(wb, sheet, SHEET_WRONG_ROI)
        delete_columns(sheet_good_roi, columns_index_wrong)
        delete_columns(sheet_wrong_roi, columns_index_good)
        # Create new sheets to write percentage calculation results
        result_above_threshold_title = "{} {} %".format(RESULT_ABOVE, str(percentage_threshold))
        result_below_threshold_title = "{} {} %".format(RESULT_BELOW, str(percentage_threshold))
        wb.create_sheet(result_above_threshold_title)
        wb.create_sheet(result_below_threshold_title)
        write_data(wb[result_above_threshold_title], columns_info_wrong)
        write_data(wb[result_below_threshold_title], columns_info_good)
    else:
        logging.critical("no column to delete...")


def normalize_selected_value(value, mean):
    return (value - mean) / mean


def calculate_mean_and_normalize_roi(wb, sheet_to_calculate, title_for_new_mean_sheet):
    normalized_sheet_title = "{} normalized".format(sheet_to_calculate)
    selected_sheet = copy_worksheet(wb, wb[sheet_to_calculate], normalized_sheet_title)
    FILTER_MAX_COL = selected_sheet.max_column
    columns_mean = {}
    logging.info("calculating means and normalizing {}...".format(sheet_to_calculate))
    for col in selected_sheet.iter_cols(min_row=FILTER_MIN_ROW, min_col=FILTER_MIN_COL, max_row=FILTER_MAX_ROW, max_col=FILTER_MAX_COL):
        sum_roi_value = 0
        number_roi_values = 0
        # Iterate a first time to calculate the mean
        for cell in col:
            if not cell.value == col[0].value:
                sum_roi_value += cell.value
                number_roi_values += 1
        mean = (sum_roi_value / number_roi_values)
        columns_mean.update({col[0].value: mean})
        # Iterate a second time to normalized
        for cell in col:
            if not cell.value == col[0].value:
                cell.value = normalize_selected_value(cell.value, mean)
    wb.create_sheet(title_for_new_mean_sheet)
    write_data(wb[title_for_new_mean_sheet], columns_mean)


def main(excel_file):
    prepare_output_files(excel_file)
    open_excel_file(excel_file)
    if not skip_background:
        subtract_background(workbook)
    filter_columns(workbook)
    if not skip_normalize:
        calculate_mean_and_normalize_roi(workbook, SHEET_GOOD_ROI, MEAN_GOOD_ROI)
        calculate_mean_and_normalize_roi(workbook, SHEET_WRONG_ROI, MEAN_WRONG_ROI)
    logging.info("writing processed data to: '{}'".format(excel_output_file))
    workbook.save(excel_output_file)
    logging.info("DONE!")
    logging.info(STARS)


#########################
## PROGRAM STARTS HERE ##
#########################

# PARSE ARGUMENTS
parser = argparse.ArgumentParser()
parser.add_argument('-sbg', '--skip-bg', dest='skip_background', action='store_true', default=False, help="skip background subtraction step")
parser.add_argument('-snz', '--skip-normalize', dest='skip_normalize', action='store_true', default=False, help="skip mean calculation and normalization steps")
parser.add_argument('-t', '--threshold', dest='threshold', type=valid_arg_threshold, default=25, help="override threshold value")
mutually_exclusive = parser.add_mutually_exclusive_group(required=True)
mutually_exclusive.add_argument('-e', '--excel', dest='excel_file', type=valid_arg_excel, help='process only one excel file')
mutually_exclusive.add_argument('-l', '--list', dest='excel_list', type=valid_arg_list, help='process a list of excel files declared in a .txt file,'
                                                                                             'only one file should be declared per line')
args = parser.parse_args()

# SET PERCENTAGE THRESHOLD VALUE (default is 25)
percentage_threshold = int(args.threshold)
logging.info("************** THRESHOLD IS SET TO: {}% **************".format(str(percentage_threshold)))

skip_background = args.skip_background
if skip_background:
    logging.info("************** BACKGROUND SUBTRACTION STEP WILL BE SKIPPED **************")

skip_normalize = args.skip_normalize
if skip_normalize:
    logging.info("************** NORMALIZATION STEP WILL BE SKIPPED **************")

# CHECK IF PROGRAM SHOULD PROCESS A EXCEL FILE OR A LIST OF EXCEL FILES
if args.excel_list is not None:
    with open(args.excel_list) as fp:
        line = fp.readline().rstrip('\n')
        while line:
            main(line)
            line = fp.readline().rstrip('\n')
elif args.excel_file is not None:
    main(args.excel_file)
