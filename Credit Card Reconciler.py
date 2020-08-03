import pandas as pd
from openpyxl import load_workbook
from string import Template
import numpy as np
from csv import reader
import xlrd
import os


locationDictionary = {'Astation' : 'Aurora Store Station', 'A PATP' : 'Aurora Store Station', 'ASTATION PATP' : 'Aurora Store Station', 'ASTATION' : 'Aurora Store Station', 'Main Office' : 'Admin', 'MAIN OFFICE' : "Admin", 'TIRE CENTER' : 'Aurora Tire Center', 'AW PATP' : 'Aurora West A Stop', 'Awest PATP' : 'Aurora West A Stop', 'AWPATP' : 'Aurora West A Stop', 'AWEST PATP' : 'Aurora West A Stop', 'CLAY CENTER' : 'Clay Center', 'Dannbrog PATP' : 'Dannebrog Station', 'DANNEBORG' : 'Dannebrog Station', 'Dannbrog Station' : 'Dannebrog Station', 'DANNEROG' : 'Dannebrog Station',  'Elwood' : 'Elwood Station', 'ELWOOD' : 'Elwood Station', 'GIBBON': 'Gibbon', 'GI PATP' : 'Grand Island', 'GISLAND' : 'Grand Island', 'GISLAND PATP' : 'Grand Island', 'Gisland' : 'Grand Island', 'GI FEED MILL' : 'Grand Island Grain & Feed', "GRANT" : 'Grant', 'GRANT PATP' : 'Grant', 'HARDY' : 'Hardy', "Harvard PATP" : "Harvard", "HARVARD" : "Harvard", "Hastings PATP" : 'Hastings', "HASTINGS"  : 'Hastings', "KEARNEY": 'Kearney', "Keen PATP" : "Keene", 'KEEN PATP' : "Keene", 'MINDEN' : 'Minden', 'MINDEN PATP': "Minden", 'Minden' : 'Minden', "POCOMOKE" : 'Pocomoke', 'StPaul' : 'St Paul Station', 'ST PAUL'  : 'St Paul Station', 'ST PAUL PATP' : 'St Paul Station', 'SUPERIOR' : 'Superior', 'UPLAND' : 'Upland', 'UPLAND PATP' : 'Upland', 'YORK PATP' : 'York', 'YORK' : 'York', 'York' : 'York'}

readFile1 = r'C:\Users\rkeenan\OneDrive - Aurora Cooperative\Documents\Development\Account Reconciler\JuneExport.xlsx'
wb = load_workbook(readFile1, data_only=True)
exportSheet = wb['Export']

with open(r'C:\Users\rkeenan\OneDrive - Aurora Cooperative\Documents\Development\Account Reconciler\WriteFile1.txt', 'w') as f1:
    f1.write("TRC Number" + "\t" +	"Account Number" + "\t" + "Account Type" + "\t" + "Account Name"  + "\t" +	"Post Date" + "\t" + "Reference" + "\t" + "Additional Reference" + "\t" + "Amount" + "\t" +	"Description" + "\t" + "Type" + "\t" + "Text"  + "\t" +	"Type"  + "\t" + "Loc #"  + "\t" +	"Loc Name" + "\n")
with open(r'C:\Users\rkeenan\OneDrive - Aurora Cooperative\Documents\Development\Account Reconciler\WriteFile1.txt', 'a') as f1:    
    for row in range(2, exportSheet.max_row + 1):
        TRC_NUMBER = exportSheet['A' + str(row)].value
        ACCOUNT_NUM = exportSheet['B' + str(row)].value
        ACCOUNT_TYPE = exportSheet['C' + str(row)].value
        ACCOUNT_NAME = exportSheet['D' + str(row)].value
        POST_DATE = exportSheet['E' + str(row)].value
        REFERENCE = exportSheet['F' + str(row)].value
        ADDITIONAL_REFERENCE = exportSheet['G' + str(row)].value
        AMOUNT = exportSheet['H' + str(row)].value
        DESCRIPTION = exportSheet['I' + str(row)].value
        TYPE_1 = exportSheet['J' + str(row)].value
        TEXT = exportSheet['K' + str(row)].value
        TYPE_2 = exportSheet['L' + str(row)].value
        LOC_NUM = exportSheet['M' + str(row)].value
        LOC_NAME = exportSheet['N' + str(row)].value
        for x, y in locationDictionary.items():
            if LOC_NAME == x:
                LOC_NAME = y
                break
            else:
                LOC_NAME = LOC_NAME
        f1.write(str(TRC_NUMBER) + "\t" + str(ACCOUNT_NUM) + "\t" + str(ACCOUNT_TYPE) + "\t" + str(ACCOUNT_NAME) + "\t" + str(POST_DATE) + "\t" + str(REFERENCE) + "\t" + str(ADDITIONAL_REFERENCE) + "\t" + str(AMOUNT) + "\t" + str(DESCRIPTION) + "\t" + str(TYPE_1) + "\t" + str(TEXT) + "\t" + str(TYPE_2) + "\t" + str(LOC_NUM) + "\t" + str(LOC_NAME) + '\n')


print("WriteFile1.txt is done!")
wb.close()
print("Closing workbook.")


readFile2 = r'C:\Users\rkeenan\OneDrive - Aurora Cooperative\Documents\Development\Account Reconciler\Jun Credit Card Reconciliation to GL.xlsx'
wb2 = load_workbook(readFile2, data_only=True)
glLedger = wb2['GeneralLedgerDetailReportList']

with open(r'C:\Users\rkeenan\OneDrive - Aurora Cooperative\Documents\Development\Account Reconciler\WriteFile2.txt', 'w') as f2:
    f2.write("GL" + "\t" +	"PC" + "\t" + "Source Name" + "\t" + 'Account' + "\t" +	'Black Column' + "\t" +	'Name' + "\t" +	'CM' + "\t" + 'Loc'  + "\t" + 'Date' + "\t" + 'Ticket' + "\t" +	'Type' + "\t" +	'Debit' + "\t" + 'Credit' + "\t" + 'Qty' + "\t" + 'Running Balance' + "\t" + 'Source Description' + "\t" + 'Location Name' + "\t" + 'Import Cleared'  + "\t" + 'Comments' + '\n')
with open(r'C:\Users\rkeenan\OneDrive - Aurora Cooperative\Documents\Development\Account Reconciler\WriteFile2.txt', 'a') as f2:    
    for row in range(833, glLedger.max_row + 1):
        GL = glLedger['A' + str(row)].value
        PC = glLedger['B' + str(row)].value
        SOURCE_NAME = glLedger['C' + str(row)].value
        ACCOUNT = glLedger['D' + str(row)].value
        BLANK = glLedger['E' + str(row)].value
        NAME = glLedger['F' + str(row)].value
        CM = glLedger['G' + str(row)].value
        LOC = glLedger['H' + str(row)].value
        DATE = glLedger['I' + str(row)].value
        TICKET = glLedger['J' + str(row)].value
        TYPE = glLedger['K' + str(row)].value
        DEBIT = glLedger['L' + str(row)].value
        CREDIT = glLedger['M' + str(row)].value
        QTY = glLedger['N' + str(row)].value
        RUNNING_BALANCE = glLedger['O' + str(row)].value
        SOURCE_DESCRIPTION = glLedger['P' + str(row)].value
        LOCATION_NAME = glLedger['Q' + str(row)].value
        IMPORT_CLEARED = glLedger['R' + str(row)].value
        COMMENTS = glLedger['S' + str(row)].value
        f2.write(str(GL) + "\t" + str(PC) + "\t" + str(SOURCE_NAME) + "\t" + str(ACCOUNT) + "\t" + str(BLANK) + "\t" + str(NAME) + "\t" + str(CM) + "\t" + str(LOC) + "\t" + str(DATE) + "\t" + str(TICKET) + "\t" + str(TYPE) + "\t" + str(DEBIT) + "\t" + str(CREDIT) + "\t" + str(QTY) + "\t" + str(RUNNING_BALANCE) + "\t" + str(SOURCE_DESCRIPTION) + "\t" + str(LOCATION_NAME) + "\t" + str(IMPORT_CLEARED) + "\t" + str(COMMENTS) + '\n')

df1 = pd.read_csv(r'C:\Users\rkeenan\OneDrive - Aurora Cooperative\Documents\Development\Account Reconciler\WriteFile1.txt', engine="python", sep='\t')
df2 = pd.read_csv(r'C:\Users\rkeenan\OneDrive - Aurora Cooperative\Documents\Development\Account Reconciler\WriteFile2.txt', engine="python", sep='\t')

INNER_JOIN = pd.merge(df1, df2, how="inner", left_on=['Loc Name', 'Amount'], right_on=['Location Name', 'Debit'])

print(INNER_JOIN)

INNER_JOIN.to_csv(r'C:\Users\rkeenan\OneDrive - Aurora Cooperative\Documents\Development\Account Reconciler\WriteFile3.txt', sep='\t')