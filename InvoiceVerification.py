import pandas as pd
import numpy as np


#Tester Files
# mieTrakInvoice_df = pd.read_excel('MTtest.xlsx', index_col=None) 
# qBInv_df = pd.read_excel('QBtest.xlsx', index_col=None, skiprows=3)

#Actual Files
mieTrakInvoice_df = pd.read_excel('Nov2022Invoices.xlsx', index_col = None) 
qBInv_df = pd.read_excel('QB_Nov_2022.xlsx', index_col = None, skiprows=3)

#Changing data types to match index
mieTrakInvoice_df['Invoice FK'] = mieTrakInvoice_df['Invoice FK'].astype(str)
qBInv_df['Num'] = qBInv_df['Num'].astype(str)

#Debug datatypes
# print(mieTrakInvoice_df.dtypes)
# print(qBInv_df.dtypes)

#Debug DataFrames/Tables
# print(mieTrakInvoice_df)
# print(qBInv_df.head(n=10))

mieTrakInvoice_df


#Set Excel writer parameters
writer = pd.ExcelWriter('testOutput.xlsx', engine='xlsxwriter', datetime_format='mm/dd/yyyy')

#Merging both files into one
mergeDF = pd.merge(qBInv_df, mieTrakInvoice_df , left_on='Num', right_on='Invoice FK', how='left')
mergeDF.to_excel(writer, sheet_name='Merged Data')

# print(mergeDF)



# Making QB Data and MT Data sheets from Stuart's specs
qbData_df = qBInv_df[['Type', 'Date', 'Num', 'Memo', 'P. O. #', 'Name', 'Item', 'Qty' , 'Sales Price' , 'Amount']]
mtData_df = mieTrakInvoice_df[['Invoice FK', 'Create Date', 'ItemFK', 'Ext Price MT']]

qBInv_df.to_excel(writer, sheet_name='QB Data')
mtData_df.to_excel(writer, sheet_name='MT Data')
# df2.to_excel(writer, sheet_name='Sheetb')
# df3.to_excel(writer, sheet_name='Sheetc')

worksheetMergData = writer.sheets['Merged Data']
worksheetQbData = writer.sheets['QB Data']
worksheetMtData = writer.sheets['MT Data']


mtRenameDict = {'Invoice FK':'Num',
                'Create Date' : 'Date',
                'Ext Price MT': 'Amount'}
mtData_df.rename(columns = mtRenameDict, inplace=True)



mtQbAlt_df = pd.concat([mtData_df, qbData_df]).sort_index(kind='merge')
mtQbAlt_df = mtQbAlt_df.sort_values(['Date', 'Num'])
# mtQbAlt_df.groupby(['Num'])


mtQbAlt_df.to_excel(writer, sheet_name='MT QB Alt')
worksheetMtQbAlt = writer.sheets['MT QB Alt']


#Formatting columns to width
colWid = len(' 4514940539 ')

worksheetMergData.set_column(1, 19, colWid)
worksheetMtQbAlt.set_column(1, 19, colWid)
worksheetQbData.set_column(1, 10, colWid)
worksheetMtData.set_column(1, 10, colWid)

writer.close()


group = mergeDF['Invoice FK']['390632']
result = mergeDF.groupby(group)['Quantity'].transform('sum')
# for i in range mieTrakInvoice_df('Invoice FK')

print(result)