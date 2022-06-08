import pandas as pd
import re
from openpyxl import load_workbook
# Brittany R Proposed May 2022.xlsx
fileName = input()

# Load entire workbook
wb = load_workbook(filename = fileName)
# Get invoice sheet
sheet = wb['Invoice']
# Get the table, idk if it will always have Table_1 as the name...
table = sheet.tables['Table_1']
# Range of the table ex: B12:O40
rStr = table.ref
# Returns the numbers in a list
head, tail = re.findall('[0-9]+', rStr)
head = int(head) - 1
# Returns the letter range ex: B:O
coloumns = re.sub(r'[0-9]', '', rStr)

# Get the name at the top left corner, idk if this will change positions?
# dfcompany = pd.read_excel(fileName, header=0, usecols='B')
# name = dfcompany.columns.tolist()

# Before manipulating the ranges of the dataframe iterate through the coloumn names search for name of company person.
dfcompany1 = pd.read_excel(fileName)
UnnamedList = dfcompany1.columns
for name in UnnamedList:
    if 'Unnamed' not in name:
        compName = name
        break
# Create dataframe, drop horizontal and vertical rows that have nothing in them
df = pd.read_excel(fileName, header=head, usecols=coloumns)
df.dropna(how='all', axis=1, inplace=True)
df.dropna(thresh=3, inplace=True)

out_csv = pd.DataFrame()
csvNames = ['*InvoiceNo', '*Customer', '*InvoiceDate', '*DueDate', 'Terms', 'Location', 'Memo', 'Item(Product/Service)', 'ItemDescription', 'ItemQuantity', 'ItemRate', '*ItemAmount', 'ItemTaxAmount']
invoiceNum = input()
numRows = df.shape[0]
invoiceList = [int(invoiceNum) + i for i in range(numRows)]
date = input()

out_csv['*InvoiceNo'] = invoiceList
out_csv['*Customer'] = df['Parent (First Name, Last Name)']
out_csv[['*InvoiceDate', '*DueDate']] = date
out_csv['Terms'] = 'Due on Receipt'
out_csv[['Location', 'Memo']] = " "
out_csv['Item(Product/Service)'] = df['Services']
out_csv['ItemDescription'] = df['Student (First Name, Last Name)'] + ' ' + df['Services'] + ' with ' + compName + '; dates of service: ' + df['Regular Session Dates'] + ' - ' + df['Length of Sessions'] + ' sessions'
out_csv['ItemQuantity'] = df['Hours'].astype(int)
out_csv['ItemRate'] = df['Column1'].astype(int)
out_csv['*ItemAmount'] = out_csv['ItemQuantity'] * out_csv['ItemRate']
out_csv['*ItemAmount'] = out_csv['*ItemAmount'].astype(int)
out_csv['ItemTaxAmount'] = 0
out_csv = out_csv[csvNames]
csvFileName = fileName.split('.')[0] + '.csv'
out_csv.to_csv(csvFileName, index=False)