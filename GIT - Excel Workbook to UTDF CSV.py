# -*- coding: utf-8 -*-
"""
Created on Tue Feb 25 14:24:05 2020

@author: 20JCC
"""
import os, csv
import pandas as pd
import numpy as np
import decimal
from decimal import Decimal

path = input('Enter directory where volume forecast spreadsheet is located: ')
os.chdir(path)
# file_name = input('Enter file name: ')
file_name = 'Volume Forecasting'
rounding5 = input("Would you like to round traffic volumes to the nearest 5? PLEASE TYPE 'yes' OR 'no': ")
def Round5(x, base=5):
    return base * round(x/base)
# Using Pandas to find list of sheet names
xcel = pd.ExcelFile(file_name + '.xlsx')
# list of sheet names
xcel.sheet_names
line1 = ['[Lanes]']
line2 = ['Lane Group Data']
line3 = ['RECORDNAME', 'INTID', 'EBL', 'EBT', 'EBR', 'WBL', 'WBT', 'WBR', 'NBL', 'NBT', 'NBR', 'SBL', 'SBT', 'SBR']
# =============================================================================
#  specific to how I name tabs
#  may need to modify search_list for each project
# =============================================================================
search_list = ['AM Exist', 'PM Exist', 'SAT Exist',
               'AM Bgd', 'PM Bgd', 'SAT Bgd',
               'AM FB 2036 - do nothing', 'PM FB 2036 - do nothing', 'SAT FB 2036 - do nothing',
               'AM TF 2036 - do nothing', 'PM TF 2036 - do nothing', 'SAT TF 2036 - do nothing',
               'AM FB 2036 - interchange', 'PM FB 2036 - interchange', 'SAT FB 2036 - interchange',
               'AM TF 2036 - interchange', 'PM TF 2036 - interchange', 'SAT TF 2036 - interchange',
               'AM Site Net - interchange', 'PM Site Net - interchange', 'SAT Site Net - interchange',
               'AM Site Net - do nothing', 'PM Site Net - do nothing', 'SAT Site Net - do nothing']       
desire_tabs = []
# exact match
for desire in search_list:
    for i in xcel.sheet_names:
        if i == desire:
            desire_tabs.append(i)
scenario = {}
for j in desire_tabs:
    scenario[j] = {}
    df = xcel.parse(j)
    # find all rows of # and puts into array. find all cols of # and puts into array
    mask = df.applymap(lambda x: '#' in str(x))
    # https://stackoverflow.com/questions/48016629/get-column-and-row-index-for-highest-value-in-dataframe-pandas
    row,col = np.where(mask.values == True)
    for i in range(len(row)):
        int_id = int(df.iloc[row[i], col[i] + 1])
        scenario[j][int_id] = {}
# =============================================================================
#         strange rounding behaviour https://stackoverflow.com/questions/10825926/python-3-x-rounding-behavior
# =============================================================================
        if rounding5 == 'no':
            scenario[j][int_id]['SBL'] = Decimal(df.iloc[row[i] + 5, col[i] + 4]).quantize(Decimal('1'),rounding = decimal.ROUND_HALF_UP)        
            scenario[j][int_id]['SBT'] = Decimal(df.iloc[row[i] + 5, col[i] + 3]).quantize(Decimal('1'),rounding = decimal.ROUND_HALF_UP)
            scenario[j][int_id]['SBR'] = Decimal(df.iloc[row[i] + 5, col[i] + 2]).quantize(Decimal('1'),rounding = decimal.ROUND_HALF_UP)
            scenario[j][int_id]['NBL'] = Decimal(df.iloc[row[i] + 6, col[i] + 5]).quantize(Decimal('1'),rounding = decimal.ROUND_HALF_UP)
            scenario[j][int_id]['NBT'] = Decimal(df.iloc[row[i] + 6, col[i] + 6]).quantize(Decimal('1'),rounding = decimal.ROUND_HALF_UP)
            scenario[j][int_id]['NBR'] = Decimal(df.iloc[row[i] + 6, col[i] + 7]).quantize(Decimal('1'),rounding = decimal.ROUND_HALF_UP)
            scenario[j][int_id]['EBL'] = Decimal(df.iloc[row[i] + 6, col[i] + 4]).quantize(Decimal('1'),rounding = decimal.ROUND_HALF_UP)
            scenario[j][int_id]['EBT'] = Decimal(df.iloc[row[i] + 7, col[i] + 4]).quantize(Decimal('1'),rounding = decimal.ROUND_HALF_UP)
            scenario[j][int_id]['EBR'] = Decimal(df.iloc[row[i] + 8, col[i] + 4]).quantize(Decimal('1'),rounding = decimal.ROUND_HALF_UP)
            scenario[j][int_id]['WBL'] = Decimal(df.iloc[row[i] + 5, col[i] + 5]).quantize(Decimal('1'),rounding = decimal.ROUND_HALF_UP)
            scenario[j][int_id]['WBT'] = Decimal(df.iloc[row[i] + 4, col[i] + 5]).quantize(Decimal('1'),rounding = decimal.ROUND_HALF_UP)
            scenario[j][int_id]['WBR'] = Decimal(df.iloc[row[i] + 3, col[i] + 5]).quantize(Decimal('1'),rounding = decimal.ROUND_HALF_UP)
# =============================================================================
#       Round to nearest 5        
# =============================================================================
        elif rounding5 == 'yes':
            scenario[j][int_id]['SBL'] = int(Round5(df.iloc[row[i] + 5, col[i] + 4], base = 5))      
            scenario[j][int_id]['SBT'] = int(Round5(df.iloc[row[i] + 5, col[i] + 3], base = 5))
            scenario[j][int_id]['SBR'] = int(Round5(df.iloc[row[i] + 5, col[i] + 2], base = 5))
            scenario[j][int_id]['NBL'] = int(Round5(df.iloc[row[i] + 6, col[i] + 5], base = 5))
            scenario[j][int_id]['NBT'] = int(Round5(df.iloc[row[i] + 6, col[i] + 6], base = 5))
            scenario[j][int_id]['NBR'] = int(Round5(df.iloc[row[i] + 6, col[i] + 7], base = 5))
            scenario[j][int_id]['EBL'] = int(Round5(df.iloc[row[i] + 6, col[i] + 4], base = 5))
            scenario[j][int_id]['EBT'] = int(Round5(df.iloc[row[i] + 7, col[i] + 4], base = 5))
            scenario[j][int_id]['EBR'] = int(Round5(df.iloc[row[i] + 8, col[i] + 4], base = 5))
            scenario[j][int_id]['WBL'] = int(Round5(df.iloc[row[i] + 5, col[i] + 5], base = 5))
            scenario[j][int_id]['WBT'] = int(Round5(df.iloc[row[i] + 4, col[i] + 5], base = 5))
            scenario[j][int_id]['WBR'] = int(Round5(df.iloc[row[i] + 3, col[i] + 5], base = 5))
    x = j.split(' ')
    char = j.split(' ',1)[1]
    # tab names without horizon years
    if len(x) == 2:
        if x[0] == 'AM':
            order = '(1)'
        elif x[0] == 'PM':
            order = '(2)'
        elif (x[0] == 'MID') or (x[0] == 'SAT'):
            order = '(3)'
        csv_name = '(' + x[1] + ')' + order + '(' + x[0] + ')'
    # tab names with horizon years
    elif len(x) == 3:
        if x[0] == 'AM':
            order = '(1)'
        elif x[0] == 'PM':
            order = '(2)'
        elif  (x[0] == 'MID') or (x[0] == 'SAT'):
            order = '(3)'
        csv_name = '(' + x[1] + ' ' + x[2] + ')' + order + '(' + x[0] + ')'
    # tab names with more than 3 parts (used for proj# 20-2801)
    elif len(x) >= 4:
        if x[0] == 'AM':
            order = '(1)'
        elif x[0] == 'PM':
            order = '(2)'
        elif  (x[0] == 'MID') or (x[0] == 'SAT'):
            order = '(3)'
        # csv_name = '(' + x[1] + ' ' + x[2] + ' ' + x[3] + ' ' + x[4] + ')' + order + '(' + x[0] + ')'    
        csv_name = '(' + char + ')' + order + '(' + x[0] + ')'    
    with open(csv_name + '.csv', 'w', newline = '') as f:
        c = csv.writer(f)
        c.writerow(line1)
        c.writerow(line2)
        c.writerow(line3)
        for INTID in scenario[j]:
            volume_line = []
            volume_line.append('Volume')
            volume_line.append(INTID)
            volume_line.append(scenario[j][INTID]['EBL'])
            volume_line.append(scenario[j][INTID]['EBT'])
            volume_line.append(scenario[j][INTID]['EBR'])
            volume_line.append(scenario[j][INTID]['WBL'])
            volume_line.append(scenario[j][INTID]['WBT'])
            volume_line.append(scenario[j][INTID]['WBR'])
            volume_line.append(scenario[j][INTID]['NBL'])
            volume_line.append(scenario[j][INTID]['NBT'])
            volume_line.append(scenario[j][INTID]['NBR'])
            volume_line.append(scenario[j][INTID]['SBL'])
            volume_line.append(scenario[j][INTID]['SBT'])
            volume_line.append(scenario[j][INTID]['SBR'])          
            c.writerow(volume_line)         
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