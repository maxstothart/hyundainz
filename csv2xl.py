import os
import csv
import pandas as pd

fin = "C:/Users/maxst/Downloads/CCF_000433.pdf"
file = open("TEMP/output.csv")
csvreader = csv.reader(file)
rows = []
for row in csvreader:
        rows.append(row)
file.close()


data= pd.read_csv("TEMP/output.csv")
#print(data)
# Select rows from list using df.loc[df.index[index_list]]
# splitting dataframe by row index
enddata  = data.iloc[10:-6:]
print(enddata)
enddata.to_excel('TEMP/table.xlsx', index=False)
