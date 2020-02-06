# -*- coding: utf-8 -*-
"""
Created on Thu Jan  2 12:42:29 2020

@author: 20JCC
"""
path = input('Enter directory where zoned reports are located: ')

import re, os, glob, csv
import pandas as pd
from pandas import DataFrame
from collections import namedtuple
# %% FUNCTIONS
def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        return False
def Sub_DataFrame(df):
    subdf = df.iloc[5:, 2:]                                                     # extracting sub dataframe, extracting row 5 and on because 'Sign Control' is in different rows for TWSC and AWSC
    subdf_ind = []
    for string in list(df.iloc[5:, 0]):
        subdf_ind.append(string.strip())                                        # stripping leading and trailing spaces  
    subdf.columns = df.iloc[4, 2:]
    subdf.index = subdf_ind
    return subdf
def is_TWSC(df):
    df = Sub_DataFrame(df)
    if df.loc['Sign Control'].str.count('Stop').sum() <= 2:                     # location of 'Sign control' 
#        print (name + ' is a Two-way stopped controlled (TWSC)' )
        return True
    elif df.loc['Sign Control'].str.count('Stop').sum() == 3 or 4:
#        print (name + ' is All-way stopped controlled (AWSC)')
        return False
def Stopped_Approach(df):
    data_df = Sub_DataFrame(df)
    # from https://stackoverflow.com/questions/41289074/find-column-names-when-row-element-meets-a-criteria-pandas
    stopped_approach = (data_df.loc['Sign Control'] == 'Stop')[lambda x: x].index.tolist()    
    for g in range(len(stopped_approach)):
        if stopped_approach[g][0] == 'W':
            stopped_approach[g] = 'W'
        elif stopped_approach[g][0] == 'E':
            stopped_approach[g] = 'E'
        elif stopped_approach[g][0] == 'N':
            stopped_approach[g] = 'N'
        elif stopped_approach[g][0] == 'S':
            stopped_approach[g] = 'S'
    return stopped_approach
def Stopped_Movements(df, stopped_approach):                                    # finds stop movements. I believe this covers all scenarios
    pseudo_stpd_mvmt = []
    stopped_movement = []
    data_df = Sub_DataFrame(df)
    lane_grp = ((data_df.loc['Lanes']!= "") & (data_df.loc['Lanes']!= "0"))[lambda x: x].index.tolist()        # finds lane groups    
    if is_TWSC(df) == True:
        r_flare = (data_df.loc['Right turn flare (veh)'] != '')[lambda x: x].index.tolist()                    # find movements that are right turn flares.
        if 'r_flare' in locals():
            for mvmt in r_flare:
                print (mvmt + ' is a right turn flare')
            for mvmt in lane_grp:
                if mvmt[0] in stopped_approach:                                 # find stopped movements
        #            print( mvmt + ' is stopped movement')
                    pseudo_stpd_mvmt.append(mvmt) 
            for mvmt in lane_grp:                                               # stopped movments that are right turn flares are not recorded SEPARATELY in 'Direction, Lane #'
                if (mvmt in pseudo_stpd_mvmt) and (mvmt not in r_flare):
                    print (mvmt + ' is ....')
                    stopped_movement.append(mvmt)
        else:
            for mvmt in lane_grp:
                if mvmt[0] in stopped_approach:
                    stopped_movement.append(mvmt)
    elif is_TWSC(df) == False:                                                  # all way stopped approach does not have right turn flare rows? - confirmed with Uni-Fab AWSC - fake. This is a weird characteristic of HCM 2000 Unsig reporting
        for mvmt in lane_grp:
            if mvmt[0] in stopped_approach:
                stopped_movement.append(mvmt)
    return stopped_movement
#    return stopped_movement, r_flare, lane_grp
#    return stopped_movement, lane_grp
def HCM_unsig_result(df):
    df = Sub_DataFrame(df)                                                      # calls previously defined function, since the Sub_DataFrame cleans the index trailing/leading spaces   
    results = df.loc['Direction, Lane #':'Approach LOS', :]                     # extract initial sub-dataframe
    col = []
    for item in df.loc['Direction, Lane #']:                                    # get names of Direction that is not ' ', will be used as new column headers
        if item != '':
            col.append(item)  
    results = results.iloc[results.index.get_loc('Direction, Lane #')+1:\
                           results.index.get_loc('Approach LOS')+1, :len(col)]  # extract sub-dataframe based on integer index
    results.columns = col                                                       # declare new column header
    return results   
def Intersection_Summ(df):
    sub_df = Sub_DataFrame(df)
    intersection_summ = sub_df.loc['Intersection Summary':, ]
    if df.iloc[0, 0] == 'HCM Unsignalized Intersection Capacity Analysis':
        for item in intersection_summ.loc['Delay']:
            if item != '':
                intersection_delay = item           
        for item in intersection_summ.loc['Level of Service']:
            if item != '':
                intersection_LOS = item
    elif df.iloc[0, 0] == 'Lanes, Volumes, Timings':        
        indices_delay = [c for c, d in enumerate(intersection_summ.index.tolist()) if 'Intersection Signal Delay' in d] # returns indices of Intersection delay
        intersection_delay = intersection_summ.index.tolist()[indices_delay[0]].split(': ')[1]
        #test = intersection_summ.iloc[indices[0]]
        indices_LOS = [b for b, a in enumerate(intersection_summ.iloc[indices_delay[0]]) if 'Intersection LOS' in a]    # intersection LOS is on the same row as intersection delau
        intersection_LOS = intersection_summ.iloc[indices_delay[0], indices_LOS[0]].split(': ')[1]
    return intersection_delay, intersection_LOS
# unsignalized - LOS Defintions
def Unsignalized_LOS(delay):
    if delay <= 10:
        los = 'A'
    elif 10 < delay <= 15:
        los = 'B'
    elif 15 < delay <= 25:
        los = 'C'
    elif 25 < delay <= 35:
        los = 'D'
    elif 35 < delay <= 50:
        los = 'E'
    elif 50 < delay:
        los = 'F'
    return los        
# signalized - LOS Defintions
def Signalized_LOS(delay):
    if delay <= 10:
        los = 'A'
    elif 10 < delay <= 20:
        los = 'B'
    elif 20 < delay <= 35:
        los = 'C'
    elif 35 < delay <= 55:
        los = 'D'
    elif 55 < delay <= 80:
        los = 'E'
    elif 80 < delay:
        los = 'F'
    return los
def Unsig_Report_Table(result_df, stopped_result, column_name, stopped_movement):
    table_mvmt = []
    for mvmt in stopped_movement:
        if mvmt[2] == 'R':
            table_mvmt.append(mvmt[0:2] + ' right')
        elif mvmt[2] == 'T':
            table_mvmt.append(mvmt[0:2] + ' through')
        elif mvmt[2] == 'L':
            table_mvmt.append(mvmt[0:2] + ' left')
    table = result_df.T                                                         # transpose
    table = table.loc[table.index.isin(stopped_result)]                         # remove indices that are not stopped approaches 
    table = table[table.columns.intersection(column_name)]                      # removing caolumns that are not relevant results
    table = table.reindex(columns = column_name)                                # reorders columns           
    table.index = table_mvmt                                                    # replace indices   
    return table
def Sig_Report_Table(df, column_name):
    data_df = Sub_DataFrame(df)    
    lane_grp = ((data_df.loc['Lane Configurations'] != "") & (data_df.loc['Lane Configurations'] != "0"))[lambda x: x].index.tolist()        # finds lane groups
    table = data_df.T                                                           # transpose
    table = table.loc[table.index.isin(lane_grp)]                               # remove indices that are not stopped approaches 
    table = table[table.columns.intersection(column_name)]                      # removing caolumns that are not relevant results  
    table = table.reindex(columns = column_name)                                # reorders columns    
    table_mvmt = []
    for mvmt in lane_grp:
        if mvmt[2] == 'R':
            table_mvmt.append(mvmt[0:2] + ' right')
        elif mvmt[2] == 'T':
            table_mvmt.append(mvmt[0:2] + ' through')
        elif mvmt[2] == 'L':
            table_mvmt.append(mvmt[0:2] + ' left')
    table.index = table_mvmt                                                    # replace indices
    return table
def Name_Shorten(string, acrynm_dict):                                          # case sensitive
    for key in acrynm_dict.keys():
        string = string.replace(key,acrynm_dict[key])
        if len(string) > 31:
            string = string.replace(' ','')
            if len(string) > 31:
                string = string[0:31]
            else:
                pass
        else:
            pass
    return string 
# %% Read Data
#os.chdir(r"C:\Users\20JCC\Documents\Projects\Dorchester\Py")
#os.chdir(r"C:\Users\20JCC\Documents\Projects\19-1170 - New Highschool\Synchro\Exist")    
#os.chdir(r"C:\Users\20JCC\Documents\Python Scripts\Synchro Extract\Jacky")       
#create named tuple
#Intersection = namedtuple('Intersection',['name', 'peak', 'scenario', 'report'])
#os.chdir(r"C:\Users\20JCC\Documents\Projects\19-1170 - New Highschool\Synchro\Exist")
os.chdir(path)
dict_csv = {}
dict_table = {}
int_name = []
int_time = []
int_scenario = []
int_key = []
for txtfile in sorted(glob.glob("*.txt")):
    f = []                                                                      # reset list after each text file is converted to dataframe
    with open (txtfile) as csvfile:     
        csv_reader = csv.reader(csvfile, delimiter = '\t')
        for row in csv_reader:
            f.append(row)       
    data_raw = DataFrame(f).fillna('')                                                                             
    sep = []
    for index, row in data_raw.iterrows():        
        if (row[0] == 'HCM Unsignalized Intersection Capacity Analysis') or (row[0] == 'Lanes, Volumes, Timings'):
            sep.append(int(index))   
    sep.append(len(data_raw))  
    for i in range(len(sep)-1):
#        print str(sep[i]) + ' and ' + str(sep[i+1])      
        data = data_raw.iloc[sep[i]:sep[i+1], :]       
        #intersection name
        id_text = data.iloc[1, 0]
        id_num, name = id_text.split(': ')
        #intersection peak hour
        peak = data.iloc[0, 1]
        #intersection scenario
        scen = data.iloc[1, 1]
        #intersection type
        report = data.iloc[0, 0]        
#        key = Intersection(name = name, peak = peak, scenario = scen, report = report)
        key = (name, peak, scen, report)
        # dictioanry of dataframes with tuples as keys
        dict_csv[key] = data.copy()
        int_name.append(name)
        int_time.append(peak)
        int_scenario.append(scen)
        int_key.append(key)
# %% Extracting relevant results from results_df and puttin into table"""
        if report == 'HCM Unsignalized Intersection Capacity Analysis':    
            stpd_aprch = Stopped_Approach(data)
            stpd_mvmt = Stopped_Movements(data, stpd_aprch)   
            """ Translates to EB 1 EB 2 WB 1 WB 2 NB 1 NB 2 etc.tc.eetc. """
            stpd_rslt = []                                                      # this works cuz HCm 2000 Unsig limits to 2 lanes per approach
            for i in stpd_aprch:
                count = 0
                for j in stpd_mvmt:
                    if i == j[0]:
                        count += 1
                        stpd_rslt.append(j[0:2] + ' ' + str(count))
            #check that stpd_rslt and stpd_mvmt are the same length
            if len(stpd_rslt) == len(stpd_mvmt):
                print ('stpd_rslt and stpd_mvmt are the same length << AWESOME')
            else:
                print ('stpd_rslt and stpd_mvmt are NOT the same length << BAD')
            if is_TWSC(data) == True:
                table_columns = ['Volume to Capacity','Lane LOS', 'Control Delay (s)','Queue Length 95th (m)'] # for TWSC ... Control delay or approach delay   
                table_df = Unsig_Report_Table(HCM_unsig_result(data), stpd_rslt, table_columns, stpd_mvmt)
                # finds missing LOS and fills in
                for index, _ in table_df.iterrows():
                    if (table_df.loc[index, 'Lane LOS'] == '') and (table_df.loc[index, 'Control Delay (s)'] != ''):
                        table_df.loc[index,'Lane LOS'] = Unsignalized_LOS(float(table_df.loc[index, 'Control Delay (s)']))
                    else:
                        pass
                    # rounds to integer if delay is over 100
                    if float(table_df.loc[index,'Control Delay (s)']) >= 100:
                        table_df.loc[index, 'Control Delay (s)y'] = str(round(float(table_df.loc[index, 'Control Delay (s)'])))
                    else:
                        pass
                    # extracts just the number from queue and rounds to zero decimal place
                    temp = re.findall("\d+\.\d+",table_df.loc[index, 'Queue Length 95th (m)'])  # extracts just the number from queue. creates a .  From https://stackoverflow.com/questions/4703390/how-to-extract-a-floating-number-from-a-string
                    if temp:                                                                    # if temp is not empty
                        table_df.loc[index, 'Queue Length 95th (m)'] = round(float(temp[0]))    # rounds queue
                    elif not temp:                                                              # if temp is empty
                        pass
            elif is_TWSC(data) == False:
                table_columns = ['Degree Utilization, x', 'Approach LOS','Control Delay (s)']
                table_df = Unsig_Report_Table(HCM_unsig_result(data), stpd_rslt, table_columns, stpd_mvmt)    
                # finds missing LOS and fills in
                for index, _ in table_df.iterrows():
                    if (table_df.loc[index, 'Approach LOS'] == '') and (table_df.loc[index, 'Control Delay (s)'] != ''):
                        table_df.loc[index, 'Approach LOS'] = Unsignalized_LOS(float(table_df.loc[index, 'Control Delay (s)']))
                    else:
                        pass
                # AWSC does not report queue
                # intersection summary
                int_delay, int_LOS = Intersection_Summ(data)
                table_df.loc['Overall', 'Control Delay (s)'] = int_delay
                table_df.loc['Overall', 'Approach LOS'] = int_LOS
                table_df = table_df.fillna('-')
        elif report == 'Lanes, Volumes, Timings':
            table_columns = ['v/c Ratio', 'LOS', 'Total Delay', 'Queue Length 95th (m)']    # rows from textfile that we want
            table_df = Sig_Report_Table(data, table_columns)
            for index, _ in table_df.iterrows():
                if (table_df.loc[index, 'LOS'] == '') and (table_df.loc[index, 'Total Delay'] != ''):
                    table_df.loc[index, 'LOS'] = Signalized_LOS(float(table_df.loc[index, 'Total Delay']))
                else:
                    pass
                # rounds to integer if delay is over 100
                if float(table_df.loc[index,'Total Delay']) >= 100:
                    table_df.loc[index, 'Total Delay'] = str(round(float(table_df.loc[index, 'Total Delay'])))
                else:
                    pass
                # extracts just the number from queue and rounds to zero decimal place
                temp = re.findall("\d+\.\d+", table_df.loc[index, 'Queue Length 95th (m)']) # extracts just the number from queue. creates a .  From https://stackoverflow.com/questions/4703390/how-to-extract-a-floating-number-from-a-string
                if temp:                                                                    # if temp is not empty
                    table_df.loc[index, 'Queue Length 95th (m)'] = round(float(temp[0]))    # rounds queue
                elif not temp:                                                              # if temp is empty
                    pass
                int_delay, int_LOS = Intersection_Summ(data)
                table_df.loc['Overall', 'Total Delay'] = int_delay
                table_df.loc['Overall', 'LOS'] = int_LOS
                table_df = table_df.fillna('-')
        dict_table[key] = table_df.copy()
# %% Write to Excel https://stackoverflow.com/questions/32957441/putting-many-python-pandas-dataframes-to-one-excel-worksheet
int_name = list(set(int_name))
int_time = list(set(int_time))
int_scenario = list(set(int_scenario))
int_key = list(set(int_key))
dict_acrynm = {'Highway': 'Hwy',
               'Road': 'Rd',
               'Boulevard':'Blvd',
               'Street': 'St',
               'Lane': 'Ln',
               'Avenue': 'Ave',
               'Circle': 'Circ',
               'Court': 'Ct',
               'Crescent': 'Cres',
               'Drive': 'Dr',
               'Expressway': 'Exp',
               'Freeway': 'Fwy',
               'Gardens': 'Gdns',
               'Heights': 'Ht',
               'Hollow': 'Hllw',
               'Wonder': 'Wndr',
               'Land': 'Lnd',
               'West': 'W',
               'North': 'N',
               'South': 'S',
               'East': 'E',
               'County': 'Cnty',
               '/': ''}
# shorten names
int_name_xcel = []
for i in range(len(int_name)):
    name = int_name[i]
    name = Name_Shorten(name, dict_acrynm)
    int_name_xcel.append(name)
# write to excel
writer = pd.ExcelWriter('Results.xlsx',engine='xlsxwriter')   
excel_row = [1]*len(int_name)                     
for tab, name in zip(int_name_xcel,int_name):
    workbook=writer.book
    worksheet=workbook.add_worksheet(tab)                                       # create tab names from the shortened names
    writer.sheets[tab] = worksheet
    worksheet.write(0, 0, name)                                                 # writes actual name in cell A1
for key, df in dict_table.items():
    index = int_name.index(key[0])
    v = Name_Shorten(key[0],dict_acrynm)
    temp = int_name_xcel.index(v)
    df.to_excel(writer,sheet_name = int_name_xcel[temp],startrow=excel_row[index], startcol=1) 
    worksheet = writer.sheets[v]
    worksheet.write(excel_row[index]+1,0, key[1] + ' // ' + key[2])             # https://stackoverflow.com/questions/43537598/write-strings-text-and-pandas-dataframe-to-excel
    excel_row[index] += len(df.index) +1                                        # add to its corresponding 1 in excel_row
writer.save()
print('\nYour results have been printed to an Excel File')
input("Press 'Enter' key to exit")