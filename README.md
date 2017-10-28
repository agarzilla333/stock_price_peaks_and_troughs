### Stock Price Peaks and Troughs
This script gets the peaks and troughs, for a given percentage change, from a time series in an excel spread sheet.
The excel sheet should have two columns w/o column labels, column A is the date (i.e. 11/16/16), column B is the price (i.e. 746.48).

#### To use this script, from a command line:
* make sure to type below in from the directory where you have the scirpt
* python find_peaks_troughs.py --PERCENTAGE_CHANGE 7.0 --INPUT_PATH "./amazonstockprice.xlsx" --OUTPUT_PATH "./amazonstockprice_output.xlsx"
* data will be taken from the INPUT_PATH and output to OUTPUT_PATH
* amazonstockprice.xlsx and amazonstockprice_output.xlsx are examples of input and output files respectively

