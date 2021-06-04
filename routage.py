import xlsxwriter
from openpyxl import load_workbook
rootBook = xlsxwriter.Workbook('fileGenerated/roote.xlsx')
val = 'SRO-31-206-293'
pdsBook = load_workbook('fileGenerated/PDS.xlsx')
pdsSheets = pdsBook.sheetnames
for sh in pdsSheets:
    sheet = pdsBook[sh]
    value = sheet.cell(row=1,column=1).value
    if str(value) == val:
        print(sh)
    else:
        pass

