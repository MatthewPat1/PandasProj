import pandas as pd

dfcompany = pd.read_excel('Brittany R Proposed May 2022.xlsx', header=0, usecols='B')
name = dfcompany.columns.tolist()

# read excel file start dataframe at row 12, make coloumns from C -> O
df = pd.read_excel('Brittany R Proposed May 2022.xlsx', header=11, usecols='C:O')

# Get coloumn names into list
colNames = df.columns.tolist()

# Remove emopty coloumns and rows
df.dropna(how='all', axis=1, inplace=True)
df.dropna(thresh=2, inplace=True)

# Create output dataframe
out_csv = pd.DataFrame()
csvNames = ['*InvoiceNo', '*Customer', '*InvoiceDate', '*DueDate', 'Terms', 'Location', 'Memo', 'Item(Product/Service)', 'ItemDescription', 'ItemQuantity', 'ItemRate', '*ItemAmount', 'ItemTaxAmount']
#input for invoice number, get number of rows, make list of that size iterating invoice number by one
date = input()
invoiceNum = input()
numRows = df.shape[0]
invoiceList = [ int(invoiceNum) + i for i in range(numRows)]

# Adding to dataframe
out_csv['*InvoiceNo'] = invoiceList
out_csv['*Customer'] = df['Parent (First Name, Last Name)']
out_csv[['*InvoiceDate', '*DueDate']] = date
out_csv['Terms'] = 'Due on Receipt'
out_csv[['Location', 'Memo']] = " "
out_csv['Item(Product/Service)'] = df['Services']
out_csv['ItemDescription'] = df['Student (First Name, Last Name)'] + ' ' + df['Services'] + ' with ' + name + ';dates of service: ' + df['Regular Session Dates'] + ' - ' + df['Length of Sessions'] + 'sessions'
out_csv['ItemQuantity'] = df['Hours']
out_csv['ItemRate'] = df['Column1']
out_csv['*ItemAmount'] = out_csv['ItemQuantity'] * out_csv['ItemRate']
out_csv['ItemTaxAmount'] = 0
# Rearrange to make sure they are in the correct coloumn position
out_csv = out_csv[csvNames]
# Make dataframe a csv file in current directory
out_csv.to_csv('out.csv')
