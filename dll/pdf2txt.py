#setup
import os
import sys
import time
from pdf2image import convert_from_path, convert_from_bytes
import pytesseract
from PIL import Image
from openpyxl import load_workbook
from datetime import date
import pandas as pd
fin = sys.argv[1]
wb = load_workbook('output\\clickme.xlsx')
coversheet = wb['Cover Sheet']
quote = wb['Quote']
os.system("python3 dll/pdf2csv.py "+ fin)
read_file = pd.read_csv (r'TEMP\output.csv')
read_file.to_excel (r'output\outputxlsx.xlsx', index = None, header=True)

#function definitions
def extract(string, after):
    v1 = output.split(after,1)[1]
    n=(v1.find("\n"))
    global out
    out = (v1[ 0 : n ])
    time.sleep(0.1)


#ocr
outputfile = open("TEMP/output.txt", "w")
pytesseract.pytesseract.tesseract_cmd = r'files/Tesseract-OCR/tesseract'
images = convert_from_path(fin, dpi=200, output_folder="TEMP", fmt="jpg", output_file="output", poppler_path = r"files/poppler-22.04.0/Library/bin")
output = (pytesseract.image_to_string(Image.open('TEMP/output0001-1.jpg')))


#get info
outputfile.write(output)
extract(output, "Parts quote: "); print(out); partsquote=out
extract(output, "Order Type: "); print(out); ordertype=out
extract(output, "Dealer Code: "); print(out); dealercode=out
extract(output, "Partstrader job#: "); print(out); partstraderjob=out
extract(output, "Reference Number: "); print(out); referencenumber=out
extract(output, "Model Code: "); print(out); modelcode=out
extract(output, "VIN: "); print(out); vin=out
extract(output, "Vehicle year: "); print(out); vehicleyear=out
extract(output, "Registration: "); print(out); registration=out
extract(output, "Closing Date: "); print(out); closingdate=out
extract(output, "Status: "); print(out); status=out
extract(output, "Buy total: "); print(out); buytotal=out
extract(output, "Sell total: "); print(out); selltotal=out
extract(output, "Admin comment: "); print(out); admincomment=out
#extract(output, "Address: "); print(out); address=out

#put info into spreadsheet
coversheet['G4'] = date.today()
coversheet['G5'] = partstraderjob
coversheet['G6'] = registration
coversheet['G8'] = modelcode
coversheet['G9'] = vehicleyear
coversheet['G11'] = vin
coversheet['K1'] = admincomment




#end
wb.save('output\\clickme.xlsx')
outputfile.close()
