from dbfread import DBF
from openpyxl import load_workbook

workbook = load_workbook('FCI 21-011/fCI-070.xlsx')
sheet = workbook.active
maxrow = sheet.max_row
# ####################### load the both file boite and cable in DBF format ###################################
supportTable = DBF('etiqueteInputs/21_011_068_SUPPORT_A.dbf', load=True, encoding='iso-8859-1')
pointTechTable = DBF('vonder/85_072_791_POINT_TECHNIQUE_B3.dbf', load=True, encoding='iso-8859-1')
pointlen = len(pointTechTable)
supplen = len(supportTable)
# FROM THE TECHNIC POINT
pointNom = []
pointCode = []
pointStruc = []
pointPrp = []
for p in range(0, pointlen):
    x = pointTechTable.records[p]['TYPE_STRUC']
    if str(x).startswith('CH'):
        pointStruc.append(x)
        pointNom.append(pointTechTable.records[p]['NOM'])
        pointCode.append(pointTechTable.records[p]['CODE'])
        pointPrp.append(pointTechTable.records[p]['PROPRIETAI'])

# FROM THE SUPPORT
suppAmount = []
suppAval = []
for s in range(0, supplen):
    suppAmount.append(supportTable.records[s]['AMONT'])
    suppAval.append(supportTable.records[s]['AVAL'])

nom = []
fci = []
for i in range(2, maxrow):
    nom.append(sheet.cell(row=i, column=1).value)
    fci.append(sheet.cell(row=i, column=2).value)


def checkfci(point):
    try:
        index = nom.index(point)
        return True
    except ValueError:
        return False


def getFci(point):
    index = nom.index(point)
    return fci[index]


k=maxrow+1
for p in pointNom:
    pass