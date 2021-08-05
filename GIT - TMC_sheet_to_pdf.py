# -*- coding: utf-8 -*-
"""
Spyder Editor

"""
import os
from openpyxl import load_workbook
import win32com.client
import shutil
from PyPDF2 import PdfFileMerger, PdfFileReader
""" This only works on traffic data that is in excel format from
    PortableStudies.com """
""" Make sure a folder is created to store the excel TMCs. this folder must be named 'Spreadsheets' """
proj_path = input('Enter directory where "TMC" folder is: ')
# %% Script
""" File paths """
path = proj_path+'\\Spreadsheets'                                               # folder where AM/PM TMC counts are
temp_folder = proj_path+'\\TMC_to_Report'                                       # temporary folder
pdf_folder = proj_path +'\\PDF'                                                 # making new PDF folder
# delete folder if exist
if os.path.exists(temp_folder):
    shutil.rmtree(temp_folder)
os.mkdir(temp_folder)                                                           # creates temporary folder for python algorithm
os.chdir(path)
""" This only works if path only has excel files """
l = os.listdir(path)                                                            # list of files in directoy with extensions
li = [x.split('.')[0] for x in l]                                               # file extension taken out
temp_ext = ' - Report Rdy.xlsx'                                                 # ' - Report Rdy.xlsx' to distinguish that Bank1 and Bank4 has changed
for tmc_files in li:
    print (tmc_files + ': Bank1 & Bank4 headings converted')
    ## from https://stackoverflow.com/questions/18849535/how-to-write-update-data-into-cells-of-existing-xlsx-workbook-using-xlsxwriter-i
    xfile = load_workbook(path + '\\' + tmc_files +'.xlsx')
    sheet = xfile['Bank1']
    sheet['A1'] = 'Passenger Car Counts'
    sheet = xfile['Bank4']
    sheet['A1'] = 'Truck Counts'
    xfile.save(temp_folder +'\\'+ tmc_files + temp_ext)
# =============================================================================
#                               SAVE EXCEL TO PDF 
# =============================================================================
""" Cannot open Excel while script below is running """
l_new = os.listdir(temp_folder)
li_new = [x.split(temp_ext)[0] for x in l_new]
os.mkdir(pdf_folder)
o = win32com.client.Dispatch('Excel.Application')
o.Visible = True                                                                # makes excel application visible
for files in li_new:
    print (files + ': Converted to PDF')
    wb = o.Workbooks.Open(temp_folder +'\\'+ files + temp_ext)
    ws_name = ['Summary','Bank1', 'Bank4']
    pdf_path = pdf_folder +'\\' + files +'.pdf'
# =============================================================================
# # fits worksheets into layout as described in Excel
# # from https://stackoverflow.com/questions/16683376/print-chosen-worksheets-in-excel-files-to-pdf-in-python/36453346
# =============================================================================
    for index in ws_name:
        ws = wb.Worksheets[index]
        ws.PageSetup.Zoom = False
        ws.PageSetup.FitToPagesTall = 1
        ws.PageSetup.FitToPagesWide = 1
    wb.WorkSheets(ws_name).Select()
    wb.ActiveSheet.ExportAsFixedFormat(0,pdf_path)
    wb.Close(False)
o.Quit
wb = None
o = None
shutil.rmtree(temp_folder)                                                      # deletes 'TMC_to_Report' folder
# =============================================================================
#                               Merge PDFS 
# =============================================================================
l_pdf = os.listdir(pdf_folder)
os.chdir(pdf_folder)                                                            # need to change working directory
merger = PdfFileMerger()
for pdf in l_pdf:
    merger.append(PdfFileReader(pdf,'rb'))
merger.write("Traffic data.pdf")
merger.close()
print ('\nTurning movement count PDFs have been combined :)')
""" The code is broken at line 56. apparently its common ... now what """
""" December 18, 2019, it works again wtf ?? """
""" need to make python stop using pdf_folder directory"""
# %%
print('\nYour results have been printed to an Excel File')
input("Press 'Enter' key to exit")
