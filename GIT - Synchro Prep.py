# -*- coding: utf-8 -*-
"""
Created on Thu Feb 13 16:57:04 2020

@author: 20JCC
"""
import os
path = input('Enter directory where Prep documents are: ')
os.chdir(path)
import csv
import shutil
#%% load report headers then copy sample and edit // text file is ordered as follows: Filename \t peak period \t Scenario \t Analyst \ Project # // 1.5 hrs on coding
f = []
scen_list = []
with open ('Report Headers.txt') as csvfile:
    csv_reader = csv.reader(csvfile, delimiter = '\t')
    for row in csv_reader:
        f.append(row)
        shutil.copyfile(path + '\\Template Sim File.txt', path + '\\' + row[0] + '.sim')       # open template and rewrite to new file name
        with open (row[0] + '.sim','r') as file:                                # open newly created file
            filedata = file.read()
# =============================================================================
# for now, comment out headers and rows that are unused 
# =============================================================================
        filedata = filedata.replace('Headers1=','Headers1=' + row[1])           # write to top right of Synchro headers
        filedata = filedata.replace('Headers3=','Headers3=' + row[2])           # write to bottom right of Synchro headers
        # filedata = filedata.replace('Headers4=','Headers4=' + row[3])           # write to top left of Synchro footers
        # filedata = filedata.replace('Headers6=','Headers6=' + row[4])           # write to bottom left of Synchro footers
#        filedata = filedata.replace('FontName0=Arial Narrow','FontName0=Calibri')    # change font Arial Narrow is default
#        filedata = filedata.replace('FontName1=Arial','FontName1=Calibri')           # change font Arial is default
        scen_list.append(row[2])
        with open (row[0] + '.sim','w') as file:                                # write data into newly copied .sim file
            file.write(filedata)
scen_list = list(set(scen_list))
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