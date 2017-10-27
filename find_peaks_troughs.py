import argparse
from time import sleep
from datetime import datetime, date
import xlrd
import xlsxwriter
import pprint

#####################
### Example Usage ###
#####################
# To use this script, from a command line, type something similar to what is below,
# make sure to type this in from the directory where you have the scirpt
# python HiLo_Ranges.py --PERCENTAGE_CHANGE 7.0 --INPUT_PATH "Your_Path/amazonstockprice.xlsx" --OUTPUT_PATH "Your_Path/amazonstockprice_output.xlsx"
# data will be taken from the input INPUT_PATH and output to OUTPUT_PATH, make sure the input file is 2 columns
# date, close_price  **** DO NOT include any titles or labels in the input file
# example: results_dict = { start_date: [start_date, end_date, start_price, end_price, percent_chg, number_days, index_end_date], ...}
### example: data = [date, cur_close_price]
if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument(
        '--PERCENTAGE_CHANGE',
        required=True,
        help='Please type in the % chg as an as decimal (i.e. 5 = 5%, 7 = 7%, etc.)')
    parser.add_argument(
        '--INPUT_PATH',
        required=True,
        help='Please type a string for the path where your data is located (i.e. "Your_Path/amazonstockprice.xlsx" )')
    parser.add_argument(
        '--OUTPUT_PATH',
        required=True,
        help='Please type a string for the path where to output (i.e. "Your_Path/amazonstockprice_output.xlsx" )')

    args = parser.parse_args()

    per_chg = float(args.PERCENTAGE_CHANGE)

    book = xlrd.open_workbook(args.INPUT_PATH)
    sh = book.sheet_by_index(0)
    data = []
    date_format = "%Y/%m/%d"
    for rx in range(sh.nrows):
        py_date = datetime(
            *
            xlrd.xldate_as_tuple(
                sh.row(rx)[0].value,
                book.datemode))
        data.insert(0, [py_date, sh.row(rx)[1].value])

    results_dict = {}
    length = len(data)
    i = 0
    end_leg = False
    while i < length:
        j = i + 1
        while j < length:
            if j == length - 1:
                end_leg = True
            # calc percentage change
            percent_chg = ((data[j][1] - data[i][1]) / data[i][1]) * 100
            if percent_chg > per_chg or percent_chg < -per_chg:
                if results_dict.get(data[i][0].date()):
                    if (percent_chg > per_chg and percent_chg > results_dict[data[i][0].date()][4] and results_dict[data[i][0].date()][4] > 0) or (
                            percent_chg < -per_chg and percent_chg < results_dict[data[i][0].date()][4] and results_dict[data[i][0].date()][4] < 0):
                        # we record a new high or low for this range because
                        # we're trending
                        results_dict[data[i][0].date()] = [data[i][0].date(), data[j][0].date(
                        ), data[i][1], data[j][1], percent_chg, (data[j][0] - data[i][0]).days, j]
                    elif (percent_chg > per_chg and percent_chg > results_dict[data[i][0].date()][4] and results_dict[data[i][0].date()][4] < 0) or (percent_chg < -per_chg and percent_chg < results_dict[data[i][0].date()][4] and results_dict[data[i][0].date()][4] > 0):
                        # skip to the previous peak/trough, to make sure and
                        # record the abrupt move in opposite direction of trend
                        i = results_dict[data[i][0].date()][6] - 1
                        break
                    # chk drawback or bounce > per_chg
                    elif abs(((data[j][1] - results_dict[data[i][0].date()][3]) / results_dict[data[i][0].date()][3]) * 100) > per_chg:
                        # skip to the previous peak
                        i = results_dict[data[i][0].date()][6] - 1
                        break
                else:  # this is the case in which it's not in the results_dict meaning have not visited price
                    results_dict[data[i][0].date()] = [data[i][0].date(), data[j][0].date(
                    ), data[i][1], data[j][1], percent_chg, (data[j][0] - data[i][0]).days, j]
            # case looking for drawbacks that percentage change doesn't meet
            # the per_chg% mark
            elif results_dict.get(data[i][0].date()):
                if abs(((data[j][1] -
                         results_dict[data[i][0].date()][3]) /
                        results_dict[data[i][0].date()][3]) *
                       100) > per_chg:  # result in which we draw back per_chg%
                    # skip to the previous peak/trough
                    i = results_dict[data[i][0].date()][6] - 1
                    break
            j += 1
        if end_leg:
            break
        i += 1

    workbook = xlsxwriter.Workbook(args.OUTPUT_PATH)
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': 1})
    date_format = workbook.add_format({'num_format': 'mmmm d yyyy'})
    worksheet.write('A1', 'Start Date', bold)
    worksheet.write('B1', 'End Date', bold)
    worksheet.write('C1', 'Start Price', bold)
    worksheet.write('D1', 'End Price', bold)
    worksheet.write('E1', 'Percent Chg', bold)
    worksheet.write('F1', 'Number Days', bold)
    worksheet.write('G1', 'End Index', bold)

    row = 1
    col = 0
    for key in sorted(results_dict.iterkeys()):
        print "%s: %s" % (key, results_dict[key])
        col = 0
        for element in results_dict[key]:
            if col < 2:
                # putting in headers for sheet
                worksheet.write_datetime(row, col, element, date_format)
            else:
                worksheet.write(row, col, element)
            col += 1
        row += 1

    workbook.close()

    pprint.pprint(len(results_dict))
    print length
