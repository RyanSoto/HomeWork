import pandas as pd
import numpy as np
from datetime import datetime


#Load Tester Files
# mieTrakInvoice_df = pd.read_excel('MTtest.xlsx', index_col=None) 
# qBInv_df = pd.read_excel('QBtest.xlsx', index_col=None, skiprows=3)

#Load Actual Files
mieTrakInvoice_df = pd.read_excel('Nov2022Invoices.xlsx', index_col = None) 
qBInv_df = pd.read_excel('QB_Nov_2022.xlsx', index_col = None, skiprows=3)

#Sort data to ensure duplicates are grouped together
mieTrakInvoice_df = mieTrakInvoice_df.sort_values(by='Invoice FK', ascending=True)
qBInv_df = qBInv_df.sort_values(by='Num', ascending=True)

#Changing data types to match index
mieTrakInvoice_df['Invoice FK'] = mieTrakInvoice_df['Invoice FK'].astype(str)
qBInv_df['Num'] = qBInv_df['Num'].astype(str)

#Debug datatypes
# print(mieTrakInvoice_df.dtypes)
# print(qBInv_df.dtypes)

#Debug DataFrames/Tables
# print(mieTrakInvoice_df)
# print(qBInv_df.head(n=10))


# # Determin Duplicates
# #Mie Trak
# dup =  mieTrakInvoice_df.duplicated(subset=['Invoice FK'], keep = False)
# # dup2 =  mieTrakInvoice_df.duplicated(subset=['Purchase Order Number'], keep = False)
# #QuickBooks
# dupQb =  qBInv_df.duplicated(subset=['Num'], keep = False)
# # dup2Qb =  qBInv_df.duplicated(subset=['Purchase Order Number'], keep = False)

# #Return duplicate booleans to dataframe 
# dup_df = mieTrakInvoice_df[dup]
# InvFk_dup = pd.Series(dup_df['Invoice FK'])
# # InvFk_dup = pd.Series(dup_df['Invoice FK']).head(n=400) #smaller test data

# dupQb_df = qBInv_df[dupQb]
# Num_dup = pd.Series(dupQb_df['Num'])
# # InvFk_dup = pd.Series(dup_df['Invoice FK']).head(n=400)

# print(pd.Series(dup_df['Invoice FK']).array)
# print(InvFk_dup.index.tolist())
# print(InvFk_dup.array)

# Determin Duplicates and Return duplicate boolens to dataframe
#Mie Trak
dup =  mieTrakInvoice_df.duplicated(subset=['Invoice FK'], keep = False)
# dup2 =  mieTrakInvoice_df.duplicated(subset=['Purchase Order Number'], keep = False)
#QuickBooks
dupQb =  qBInv_df.duplicated(subset=['Num'], keep = False)
# dup2Qb =  qBInv_df.duplicated(subset=['Purchase Order Number'], keep = False)

#Return duplicate booleans to dataframe then seperate for parsing
#Mie Trak
dup_df = mieTrakInvoice_df[mieTrakInvoice_df.duplicated(subset=['Invoice FK'], keep = False)]
InvFk_dup = pd.Series(dup_df['Invoice FK'])
# InvFk_dup = pd.Series(dup_df['Invoice FK']).head(n=100) #smaller test data

#Quick Books
dupQb_df = qBInv_df[qBInv_df.duplicated(subset=['Num'], keep = False)]
Num_dup = pd.Series(dupQb_df['Num'])
# InvFk_dup = pd.Series(dup_df['Invoice FK']).head(n=400)


# Duplicate counter
# for i in InvFk_dup.array:
#     for j in InvFk_dup.index.array:
#         if temp != i:
#             temp = 0
#             counter += 1
#         temp = i

# print(counter)

#Function to parse duplicates into a dictionary for indexing
def dupParser( x , y ):

    #Set Variables for Loop
    temp = 0
    dupindex = []
    dupDict = {}
    counter = 0

    #For loop creating a Duplication dictionary {Index : { Duplicate count : Duplicate Index}}
    for i, j in zip(x , y):
        if temp != i:
            temp = 0
            counter += 1
            dupindex = []
            # print('i !=  temp')         
        temp = i
        dupindex.append(j)
        dupDict[counter] = [len(dupindex) , dupindex] # len(dupindex) can be removed here and used as needed and same with the dictionary keys actually 

    # print(counter)
    # print(dupDict)
    return dupDict

mtDupDict = dupParser(InvFk_dup.array  , InvFk_dup.index.array)
# print(mtDupDict)
qbDupDict = dupParser(Num_dup.array  , Num_dup.index.array)
# print(qbDupDict)

#Duplication consolidator: find the sum of each duplication's price and quantity

# print(mieTrakInvoice_df)


#Function to index through duplicates as they are grouped
def dupCon(dupDict):
    count = 0
    temp = 0
    group = []
    quantity = 0
    tmp = []
    newFrame = pd.DataFrame()
    # newFrame.columns = ["Invoice FK", "Create Date", "ItemFK" , "Purchase Order Number", "Quantity", "Price", "Ext Price MT", "BOM Costs", "COGS"]
    # newFrame = []
    for i , j in dupDict.items():
        if temp != i:
            group = (j[1])
            for k in group:
                count += 1
                temp2 = mieTrakInvoice_df.iloc[k] 
                # print(temp2)
                # tmp.append(temp2)
                newFrame = newFrame.append(temp2)
                # quantity = quantity + mieTrakInvoice_df.iloc[k] 
        # temp.append(i)
            
        temp = i
    # newFrame = pd.concat(tmp, ignore_index=True, axis=1)
    # print(newFrame)

    return newFrame

    
    
    # print(count)



proFrame = dupCon(mtDupDict)
# proFrame = proFrame.T.reset_index().reindex(columns=["Invoice FK", "Create Date", "ItemFK" , "Purchase Order Number", "Quantity", "Price", "Ext Price MT", "BOM Costs", "COGS"])
# proFrame = proFrame.pivot_table(index="Invoice FK", columns=["Create Date", "ItemFK" , "Purchase Order Number", "Quantity", "Price", "Ext Price MT", "BOM Costs", "COGS"] )
print(proFrame)


# if dup.all() != dup2.all():
#     print('Mismatch found')
# else:
#     print('No Mismatches')


# # Writing everything to Excel
now = datetime.now()
current_time = now.strftime("%H-%M-%S")
writer = pd.ExcelWriter('tester' + current_time + '.xlsx', engine='xlsxwriter', datetime_format='mm/dd/yyyy')
mieTrakInvoice_df.to_excel(writer, sheet_name='Raw MT Data')
qBInv_df.to_excel(writer, sheet_name='Raw QB Data')
InvFk_dup.to_excel(writer, sheet_name='dups')
dup_df.to_excel(writer, sheet_name='dupsfull')
proFrame.to_excel(writer, sheet_name='proFrame')

writer.close()