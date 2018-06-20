import xlrd
import xlwt
import os
import openpyxl
import numpy as np
import regex as re
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import numpy as np

Value = []
Month1 = []
Month2 = []
Months = []

for i in range(2001,2018): 
    filepath = "/Users/ps/Documents/Transsight/BART-Open Data Portal/Datasets/Datasets for Upload/Ridership/ridership_" + str(i)
    for filename in os.listdir(filepath):
        if filename.endswith('Entry_Exit Matrices.xls'):
            Month1.append(filename[:-24])
        
            f = "/Users/ps/Documents/Transsight/BART-Open Data Portal/Datasets/Datasets for Upload/Ridership/ridership_" + str(i)+ "/" + filename 
            ow = xlrd.open_workbook(f).sheet_by_index(1)
            cnt = 0
            rcnt = 0
            c = ow.ncols
            r = ow.nrows
            # print (c)
            # print (r) 
            for x in range (0, c+1):
                try:
                    if ow.cell_value(1, x) == "Exits":
                        cnt = cnt + 1
                        # print ("column: ", x)
                        # print("row: ", r)
                        # print(cnt)          
                        ips = ow.cell_value(rowx = r-1, colx = x)
                        Value.append(ips)
                        break
                except IndexError:
                        continue
            for y in range (0, r):
                try:
                    if ow.cell_value(r, 0) == "Entries":
                        rcnt = rcnt + 1
                        # print ("column: ", x)
                        # print("row: ", r)
                        # print(cnt)          
                        rps = ow.cell_value(rowx = r, colx = c)
                        Value.append(rps)
                        break
                except IndexError:
                        continue
        elif filename.startswith('Ridership_') and filename.endswith('.xls'):
            Month2.append(filename[10:-4])
            f = "/Users/ps/Documents/Transsight/BART-Open Data Portal/Datasets/Datasets for Upload/Ridership/ridership_" + str(i)+ "/" + filename 
            ow = xlrd.open_workbook(f).sheet_by_index(1)
            cnt = 0
            c = ow.ncols
            r = ow.nrows
            # print (c)
            # print (r) 
            for x in range (0, c+1):
                try:
                    if ow.cell_value(1, x) == "Exits":
                        cnt = cnt + 1
                        # print ("column: ", x)
                        # print("row: ", r)
                        # print(cnt)          
                        ips = ow.cell_value(rowx = r-1, colx = x)
                        Value.append(ips)
                       
                        break
                except IndexError:
                        continue
        elif filename.startswith('Ridership_') and filename.endswith('.xlsx'): 
            Month2.append(filename[10:-5])
            f = "/Users/ps/Documents/Transsight/BART-Open Data Portal/Datasets/Datasets for Upload/Ridership/ridership_" + str(i)+ "/" + filename 
            ow = xlrd.open_workbook(f).sheet_by_index(1)
            cnt = 0
            c = ow.ncols
            r = ow.nrows
            

            # print (c)
            # print (r) 
            for x in range (0, c+1):
                try:
                    if ow.cell_value(1, x) == "Exits":
                        cnt = cnt + 1
                        # print ("column: ", x)
                        # print("row: ", r)
                        # print(cnt)          
                        ips = ow.cell_value(rowx = r-1, colx = x)
                        Value.append(ips)
                       
                        break
                except IndexError:
                        continue

# //------------------------------------------------------------------------------------------------------
#  writing to excel
# //------------------------------------------------------------------------------------------------------

Months = Month1 + Month2

# string="This is a string that contains #134534 and other things"
Years = [re.findall(r'\d+',m) for m in Months];
# reduce(lambda x, y: x.extend(y), Years)
Years = sum(Years,[])
# print(Months)
# print(Value)
# print(Years)
df = pd.DataFrame({'Year': Years, 'Month': Months,'Ridership': Value})                   
writer = pd.ExcelWriter('/Users/ps/Documents/FinalSat.xlsx')
df.to_excel(writer,'Sundays',index=False)
writer.save()

# ----------------------------------------------------------------------

# import xlrd
# import xlwt
# import os
# import openpyxl
# import numpy as np
# import regex
# filename = "/Users/ps/Documents/Python/test.xlsx"
# ow = xlrd.open_workbook(filename).sheet_by_index(0)
# cnt = 0
# c = ow.ncols
# r = ow.nrows

# print (c)
# print (r) 
# for x in range (0, c+1):
#     try:
#         if ow.cell_value(2, x) == "Name":
#             cnt = cnt + 1
#             print ("column: ", x)
#             print("row: ", r)
#             print(cnt)


          
#             ips = ow.cell_value(rowx = r-1, colx = x)
#             print(ips)
#             break
#     except IndexError:
#         continue