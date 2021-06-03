from dbfread import DBF
import xlsxwriter
import datetime

# date configuration
now = datetime.datetime.now()
date = now.strftime("%d/%m/%Y")
# ################## load the both file boite and cable in DBF format ###################################
cableTable = DBF('fileGenerated/069_CABLE_OPTIQUE.dbf', load=True, encoding='iso-8859-1')
boiteTable = DBF('fileGenerated/BOITE_OPTIQUE.dbf', load=True, encoding='iso-8859-1')
# ################### declare the excel file ###########################################################
workbook = xlsxwriter.Workbook('fileGenerated/pds.xlsx')
# charge the name of all filed in tables
filedCableNam = cableTable.field_names
filedBoiteNam = boiteTable.field_names
boiteLen = len(boiteTable)
cableLen = len(cableTable)
# declare the table that i need te full
boiteCode = []
boiteCable = []
boiteCableState = []
boiteReference = []
nbf = []
cableName = []
cableOrigin = []
cableExtremity = []
cableCapacity = []

for i in range(0, cableLen):
    cableName.append(cableTable.records[i]['NOM'])
    cableOrigin.append(cableTable.records[i]['ORIGINE'])
    cableExtremity.append(cableTable.records[i]['EXTREMITE'])
    cableCapacity.append(cableTable.records[i]['CAPACITE'])
for j in range(0, boiteLen):
    boiteCode.append(boiteTable.records[j]['NOM'])
    boiteCable.append(boiteTable.records[j]['AMONT'])
    boiteCableState.append(boiteTable.records[j]['INTERCO'])
    boiteReference.append(boiteTable.records[j]['REFERENCE'])

# define the character and style of cell inside excel
bold = workbook.add_format({'bold': True, "border": 1})
bold1 = workbook.add_format({'bold': True})
border = workbook.add_format({"border": 1})
back = workbook.add_format({"bg_color": '#CD5C5C', "border": 1})
cassette = workbook.add_format({"bg_color": '#A9A9A9', "border": 1})
cell_formatCapacity = workbook.add_format({"bg_color": '#E6E6FA', "border": 1})
cell_format1 = workbook.add_format({"bg_color": 'red', "border": 1})
cell_format2 = workbook.add_format({"bg_color": 'blue', "border": 1})
cell_format3 = workbook.add_format({"bg_color": '#00FF00', "border": 1})
cell_format4 = workbook.add_format({"bg_color": 'yellow', "border": 1})
cell_format5 = workbook.add_format({"bg_color": '#BF00FF', "border": 1})
cell_format6 = workbook.add_format({"bg_color": 'white', "border": 1})
cell_format7 = workbook.add_format({"bg_color": '#FFBF00', "border": 1})
cell_format8 = workbook.add_format({"bg_color": '#828282', "border": 1})
cell_format9 = workbook.add_format({"bg_color": '#816B56', "border": 1})
cell_format10 = workbook.add_format({"bg_color": '#333333', "border": 1})
cell_format11 = workbook.add_format({"bg_color": '#00FFBF', "border": 1})
cell_format12 = workbook.add_format({"bg_color": '#FFAAD4', "border": 1})
colorList = [cell_format1, cell_format2, cell_format3, cell_format4, cell_format5, cell_format6, cell_format7,
             cell_format8, cell_format9, cell_format10, cell_format11, cell_format12]


for b in range(0, boiteLen):
    N = 1
    n = 1
    w = workbook.add_worksheet(boiteCode[b])
    # INFORMATION ABOUT BOITE
    w.write('Q1', 'Etiquette : ', border)
    w.write('R1', boiteCode[b], bold)
    w.write('Q2', 'Reference : ', border)
    w.write('R2', boiteReference[b], bold)
    w.write('Q3', 'Date de modification : ', bold)
    w.write('R3', date, bold)
    # INFORMATION OF THE HEADER
    w.write('A1', 'Entrée', bold)
    w.write('B1', 'Capacité', bold)
    w.write('C1', 'N°         ', bold)
    w.write('D1', 'N° Tube', bold)
    w.write('E1', 'N° Fibre', bold)
    w.write('F1', 'Cassette', bold)
    w.write('G1', 'Etat fibre', bold)
    w.write('H1', 'N° Fibre', bold)
    w.write('I1', 'N° Tube', bold)
    w.write('J1', 'N°       ', bold)
    w.write('K1', 'Capacité', bold)
    w.write('L1', '', bold)
    w.write('M1', 'Sortie', bold)
    w.write('N1', 'Statut', bold)
    w.write('O1', 'Client', bold)
    indexCable = cableName.index(boiteCable[b])
    cap = cableCapacity[indexCable]
    for c in range(2, cap + 2):
        w.write('A' + str(c), cableName[indexCable], bold1)
        w.write('B' + str(c), cap, cell_formatCapacity)
        if n == 12:
            w.write('C' + str(c), n)
            w.write('E' + str(c), n, colorList[n - 1])
            w.write('D' + str(c), N, colorList[N - 1])
            n = 1
            if N==12:
                N=1
            else :
                N = N+1

        else:
            w.write('C' + str(c), n)
            w.write('E' + str(c), n, colorList[n - 1])
            w.write('D' + str(c), N, colorList[N - 1])
            n = n + 1


workbook.close()
