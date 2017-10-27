 To use this script, from a command line, type something similar to what is below,
 make sure to type this in from the directory where you have the scirpt
 python HiLo_Ranges.py --PERCENTAGE_CHANGE 7.0 --INPUT_PATH "Your_Path/amazonstockprice.xlsx" --OUTPUT_PATH "Your_Path/amazonstockprice_output.xlsx"
 data will be taken from the input INPUT_PATH and output to OUTPUT_PATH, make sure the input file is 2 columns
 date, close_price  **** DO NOT include any titles or labels in the input file
 example: results_dict = { start_date: [start_date, end_date, start_price, end_price, percent_chg, number_days, index_end_date], ...}
 example: data = [date, cur_close_price]