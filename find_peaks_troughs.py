import argparse, xlrd, xlsxwriter
from datetime import datetime, date

### example: results_dict = { start_date: [start_date, end_date, start_price, end_price, percent_chg, number_days, index_end_date], ...}
### example: data = [date, cur_close_price]
class Peaks_And_Troughs(object):

    def __init__(self, INPUT_PATH, OUTPUT_PATH, per_chg):
        self.INPUT_PATH = INPUT_PATH
        self.OUTPUT_PATH = OUTPUT_PATH
        self.data = self.get_dates_and_prices()
        self.per_chg = per_chg
        self.results_dict = {}
        
    def get_dates_and_prices(self):
        # get dates in a python friendly format
        data = []
        book = xlrd.open_workbook(self.INPUT_PATH)
        sh = book.sheet_by_index(0)
        date_format = "%Y/%m/%d"
        for rx in range(sh.nrows):
            py_date = datetime(
                *
                xlrd.xldate_as_tuple(
                    sh.row(rx)[0].value,
                    book.datemode))
            data.insert(0, [py_date, sh.row(rx)[1].value])
        return data
        
    def get_peaks_troughs(self):
        length = len(self.data)
        i = 0
        end_leg = False
        while i < length:
            j = i + 1
            while j < length:
                if j == length - 1:
                    end_leg = True
                # calc percentage change
                percent_chg = ((self.data[j][1] - self.data[i][1]) / self.data[i][1]) * 100
                if percent_chg > self.per_chg or percent_chg < -self.per_chg:
                    if self.results_dict.get(self.data[i][0].date()):
                        if (percent_chg > self.per_chg \
                            and percent_chg > self.results_dict[self.data[i][0].date()][4] \
                            and self.results_dict[self.data[i][0].date()][4] > 0) \
                            or (percent_chg < -self.per_chg \
                            and percent_chg < self.results_dict[self.data[i][0].date()][4] \
                            and self.results_dict[self.data[i][0].date()][4] < 0):
                            # we record a new high or low for this range because
                            # we're trending
                            self.results_dict[self.data[i][0].date()] = [self.data[i][0].date(), self.data[j][0].date(), \
                            self.data[i][1], self.data[j][1], percent_chg, (self.data[j][0] - self.data[i][0]).days, j]
                        elif (percent_chg > self.per_chg \
                            and percent_chg > self.results_dict[self.data[i][0].date()][4] \
                            and self.results_dict[self.data[i][0].date()][4] < 0) \
                            or (percent_chg < -self.per_chg \
                            and percent_chg < self.results_dict[self.data[i][0].date()][4] \
                            and self.results_dict[self.data[i][0].date()][4] > 0):
                            # skip to the previous peak/trough, to make sure and
                            # record the abrupt move in opposite direction of trend
                            i = self.results_dict[self.data[i][0].date()][6] - 1
                            break
                        # chk drawback or bounce > self.per_chg
                        elif abs(((self.data[j][1] - self.results_dict[self.data[i][0].date()][3]) 
                            / self.results_dict[self.data[i][0].date()][3]) * 100) > self.per_chg:
                            # skip to the previous peak
                            i = self.results_dict[self.data[i][0].date()][6] - 1
                            break
                    else:  # this is the case in which it's not in the self.results_dict meaning have not visited price
                        self.results_dict[self.data[i][0].date()] = [self.data[i][0].date(), self.data[j][0].date(
                        ), self.data[i][1], self.data[j][1], percent_chg, (self.data[j][0] - self.data[i][0]).days, j]
                # case looking for drawbacks that percentage change doesn't meet
                # the self.per_chg% mark
                elif self.results_dict.get(self.data[i][0].date()):
                    if abs(((self.data[j][1] -
                            self.results_dict[self.data[i][0].date()][3]) /
                            self.results_dict[self.data[i][0].date()][3]) *
                        100) > self.per_chg:  # result in which we draw back self.per_chg%
                        # skip to the previous peak/trough
                        i = self.results_dict[self.data[i][0].date()][6] - 1
                        break
                j += 1
            if end_leg:
                break
            i += 1
    
    def write_to_excel_file(self):
        workbook = xlsxwriter.Workbook(self.OUTPUT_PATH)
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
        for key in sorted(self.results_dict.keys()):
            print("%s: %s" % (key, self.results_dict[key]))
            col = 0
            for element in self.results_dict[key]:
                if col < 2:
                    # putting in headers for sheet
                    worksheet.write_datetime(row, col, element, date_format)
                else:
                    worksheet.write(row, col, element)
                col += 1
            row += 1

        workbook.close()

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument(
        '--PERCENTAGE_CHANGE',
        required=True,
        help='Please type in the % chg as an as decimal (i.e. 5 = 5%, 7 = 7%, etc.)')
    parser.add_argument(
        '--INPUT_PATH',
        required=True,
        help='Please type a string for the path where your self.data is located (i.e. "Your_Path/amazonstockprice.xlsx" )')
    parser.add_argument(
        '--OUTPUT_PATH',
        required=True,
        help='Please type a string for the path where to output (i.e. "Your_Path/amazonstockprice_output.xlsx" )')

    args = parser.parse_args()

    pk = Peaks_And_Troughs(args.INPUT_PATH, args.OUTPUT_PATH, float(args.PERCENTAGE_CHANGE))
    pk.get_peaks_troughs()
    pk.write_to_excel_file()


    

 

    