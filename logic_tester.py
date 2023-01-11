import pandas as pd
import numpy as np
from datetime import datetime


#Load Tester Files
# mieTrakInvoice_df = pd.read_excel('MTtest.xlsx', index_col=None) 
# qBInv_df = pd.read_excel('QBtest.xlsx', index_col=None, skiprows=3)

#Load Actual Files
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


# Determin Duplicates
dup =  mieTrakInvoice_df.duplicated(subset=['Invoice FK'], keep = False)
# dup2 =  mieTrakInvoice_df.duplicated(subset=['Purchase Order Number'], keep = False)
# dups = mieTrakInvoice_df.duplicated()

dupQb =  qBInv_df.duplicated(subset=['Num'], keep = False)
# dup2Qb =  qBInv_df.duplicated(subset=['Purchase Order Number'], keep = False)


colInv = pd.Series(mieTrakInvoice_df['Quantity']).head(n=10).array
col = np.array(colInv)
dup_np = np.array(dup)

dup_df = mieTrakInvoice_df[dup]
InvFk_dup = pd.Series(dup_df['Invoice FK'])
# InvFk_dup = pd.Series(dup_df['Invoice FK']).head(n=400)


dupQb_df = qBInv_df[dupQb]
Num_dup = pd.Series(dupQb_df['Num'])
# InvFk_dup = pd.Series(dup_df['Invoice FK']).head(n=400)

# print(pd.Series(dup_df['Invoice FK']).array)
# print(InvFk_dup.index.tolist())
# print(InvFk_dup.array)

def dupParser( x , y ):

    #Set Variables for Loop
    temp = 0
    dupindex = []
    dupDict = {}
    counter = 0

    # Checks how many duplications there are
    # for i in InvFk_dup.array:
    #     for j in InvFk_dup.index.array:
    #         if temp != i:
    #             temp = 0
    #             counter0 += 1
    #         temp = i

    # print(counter0)

    #For loop creating a Duplication dictionary {Index : { Duplicate count : Duplicate Index}}
    for i, j in zip(x , y):
        if temp != i:
            temp = 0
            counter += 1
            dupindex = []
            # print('i !=  temp')         
        temp = i
        dupindex.append(j)
        dupDict[counter] = {len(dupindex) : dupindex}

    # print(counter)
    # print(dupDict)
    return dupDict

mtDupDict = dupParser(InvFk_dup.array  , InvFk_dup.index.array)
# print(mtDupDict)

qbDupDict = dupParser(Num_dup.array  , Num_dup.index.array)
print(qbDupDict)



# #Set Variables for Loop
# temp = 0
# dupindex = []
# dupCount = 0
# dupDict = {}
# counter = 0

# # Checks how many duplications there are
# for i in InvFk_dup.array:
#     for j in InvFk_dup.index.array:
#         if temp != i:
#             temp = 0
#             counter0 += 1
#         temp = i

# print(counter0)

# # For loop creating a Duplication dictionary {Index : { Duplicate count : Duplicate Index}}
# for i, j in zip(InvFk_dup.array ,InvFk_dup.index.array):
#     if temp != i:
#         temp = 0
#         counter += 1
#         dupCount = 0
#         dupindex = []
#         # print('i !=  temp') 
        
#     temp = i
#     counter += 1
#     dupindex.append(j)
#     dupDict[x] = {dupCount + counter : dupindex}
        

    

# # print(counter)

# print(dupDict)


# print(col)

# count = 0
# for i in colInv:
#     if i == 1:
#         count += 1

# print(count)

# if dup.all() != dup2.all():

#     print('Mismatch found')
    
# else:
#     print('No Mismatches')

# print(colInv)
# print(dup.head(n=305))

# now = datetime.now()
# current_time = now.strftime("%H-%M-%S")

# writer = pd.ExcelWriter('tester' + current_time + '.xlsx', engine='xlsxwriter', datetime_format='mm/dd/yyyy')

# mieTrakInvoice_df.to_excel(writer, sheet_name='Raw MT Data')
# InvFk_dup.to_excel(writer, sheet_name='dups')

# writer.close()