"""
thi file is for epessourage operation still in update and progress
"""
# import xlsxwriter
import xlsxwriter
from openpyxl import load_workbook
# load your pds file here
pds = load_workbook('fileGenerated/31_206_326pds.xlsx')
wpds = pds.sheetnames

# create the epesourege file
epesBook = xlsxwriter.Workbook('31_206_326Epes.xlsx')
wr = epesBook.add_worksheet()
print(wpds)
boiteList = sorted(wpds)
print(boiteList)
header = epesBook.add_format({'bold': True, 'border': 1, 'bg_color': '#037d50'})
border = epesBook.add_format({"border": 1})
# ################# the part of coping values from pds to new file ######################
wr.write('A1', 'cable_Origine', header)
wr.write('B1', 'tube', header)
wr.write('C1', 'fibre ', header)
wr.write('D1', 'boite', header)
wr.write('E1', 'cable_Distination', header)
wr.write('F1', 'TUBE2 ', header)
wr.write('G1', 'FIBRE2', header)
wr.write('H1', 'cassete', header)
wr.write('I1', 'TYPE', header)
b = 0
for s in boiteList:
    sheet = pds[s]
    MaxRow = sheet.max_row
    MaxCol = sheet.max_column
    b += 2
    for i in range(12, MaxRow+1):
        cable=sheet.cell(row=i, column=1).value
        wr.write('A' + str(b), cable, border)
        tube = sheet.cell(row=i, column=5).value
        wr.write('B' + str(b), tube, border)
        fibre = sheet.cell(row=i, column=6).value
        wr.write('C' + str(b), fibre, border)
        # boite
        wr.write('D' + str(b), s, border)
        cableDist = sheet.cell(row=i, column=14).value
        wr.write('E' + str(b), cableDist, border)
        tube2 = sheet.cell(row=i, column=10).value
        wr.write('F' + str(b), tube2, border)
        fibre2 = sheet.cell(row=i, column=9).value
        wr.write('G' + str(b), fibre2, border)
        cassete = sheet.cell(row=i, column=7).value
        wr.write('H' + str(b), cassete, border)
        type = sheet.cell(row=i, column=8).value
        wr.write('I' + str(b), type, border)
        b += 1

epesBook.close()

