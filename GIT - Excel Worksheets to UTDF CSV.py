# -*- coding: utf-8 -*-
"""
Created on Tue Feb 25 14:24:05 2020

@author: 20JCC
"""
""" saves preset csv tabs from volume forecasting sheet"""
import os, csv
import pandas as pd
import decimal
from decimal import Decimal
import openpyxl
path = input('Enter directory where excel sheet is: ')
os.chdir(path)
file_name = input('Enter file name: ')
# =============================================================================
# Using Pandas to find list of sheet names
# =============================================================================
xcel = pd.ExcelFile(file_name + '.xlsx')
# list of sheet names
sheets = xcel.sheet_names
# sheets after UTDF
desired_sheets = sheets[sheets.index("UTDF") + 1:]
for tabs in desired_sheets:
    wb = openpyxl.load_workbook(file_name + '.xlsx', data_only = True)
    sh = wb[tabs]
    with open(tabs + '.csv', 'w', newline = "") as f:
        c = csv.writer(f)
        for r in range(1,sh.max_row+1):
            csv_row = []
            if sh[r][0].value == 'Volume':
                for item in sh[r]:
                    if (type(item.value) == str or item.value is None):
                        csv_row.append(item.value)
                    else:
                        csv_row.append(Decimal(float(item.value)).quantize(Decimal('1'),rounding = decimal.ROUND_HALF_UP))
                c.writerow(csv_row)
            elif sh[r][0].value == 'HeavyVehicles':
                for item in sh[r]:
                    if (type(item.value) == str or item.value is None):
                        csv_row.append(item.value)
                    else:
                        csv_row.append(Decimal(float(item.value)).quantize(Decimal('1'),rounding = decimal.ROUND_HALF_UP))
                c.writerow(csv_row)
            elif sh[r][0].value == 'PHF':
                for item in sh[r]:
                    if (type(item.value) == str or item.value is None):
                        csv_row.append(item.value)
                    else:
                        csv_row.append(Decimal(float(item.value)).quantize(Decimal('1.00'),rounding = decimal.ROUND_HALF_UP))
                c.writerow(csv_row)
            else:
                for item in sh[r]:
                    csv_row.append(item.value)
                c.writerow(csv_row)        
# keeping variables causes issues when saving the spreadsheet
# https://stackoverflow.com/questions/45853595/spyder-clear-variable-explorer-along-with-variables-from-memory#:~:text=Go%20to%20the%20IPython%20console,That's%20it.
def clear_all():
    """Clears all the variables from the workspace of the spyder application."""
    gl = globals().copy()
    for var in gl:
        if var[0] == '_': continue
        if 'func' in str(globals()[var]): continue
        if 'module' in str(globals()[var]): continue

        del globals()[var]
if __name__ == "__main__":
    clear_all()