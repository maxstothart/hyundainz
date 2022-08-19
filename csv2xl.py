import os
import csv
import pandas as pd
from openpyxl import load_workbook

fin = "C:/Users/maxst/Downloads/CCF_000433.pdf"
file = open("TEMP/output.csv")
csvreader = csv.reader(file)
rows = []
for row in csvreader:
        rows.append(row)
file.close()
def celladd(in1, in2):
        v1=(sh[in1].value)
        v2=(sh[in2].value)
        global out
        v3=(v1+v2)
        out=(v3[1:])


data= pd.read_csv("TEMP/output.csv")
#print(data)
# Select rows from list using df.loc[df.index[index_list]]
# splitting dataframe by row index
enddata  = data.iloc[12:-6:]
print(enddata)
enddata.to_excel('TEMP/table.xlsx', index=False)
wb = load_workbook("TEMP\\table.xlsx", data_only=True)
sh = wb["Sheet1"]
v1=(sh["a4"].value)
v2=(sh["a5"].value)
celladd("a4", "a5")
print(out)
