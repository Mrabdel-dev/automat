from openpyxl import load_workbook

# load your pds file here
pdsFile = ''
pds = load_workbook('C3A 334.xlsx')
sheet = pds['Commandes Fermes']
wpds = pds.sheetnames

maxRow = sheet.max_row
for i in range(15, maxRow):
    val = str(sheet.cell(row=i, column=3).value)
    if val is not None:
        try:
            ind = val.index('/')
        except ValueError:
            continue
        val = val[ind + 1:] + "/" + val[0:ind]
        sheet.cell(row=i, column=3).value = val
    val1 = str(sheet.cell(row=i, column=5).value)
    if val1 is not None:
        try:
            ind = val1.index('/')
        except ValueError:
            continue
        val1 = val1[ind + 1:] + "/" + val1[0:ind]
        sheet.cell(row=i, column=5).value = val1

pds.save('C3A 334 NEW.xlsx')
