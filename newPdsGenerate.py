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
boiteCode = [] ; boiteCable = [] ; boiteCableState = [] ; boiteReference = [] ; nbf = []
cableName = [] ; cableOrigin = [] ; cableExtremity = [] ; cableCapacity = []

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
bold = workbook.add_format({'bold': True,"border":1})
border = workbook.add_format({"border":1})
back = workbook.add_format({"bg_color": '#CD5C5C' , "border":1})
cassette = workbook.add_format({"bg_color": '#A9A9A9' ,"border":1})
cell_formatCapacity = workbook.add_format({"bg_color": '#E6E6FA',"border":1})
cell_format1 = workbook.add_format({"bg_color": 'red',"border":1})
cell_format2 = workbook.add_format({"bg_color": 'blue',"border":1})
cell_format3 = workbook.add_format({"bg_color": '#00FF00',"border":1})
cell_format4 = workbook.add_format({"bg_color": 'yellow',"border":1})
cell_format5 = workbook.add_format({"bg_color": '#BF00FF',"border":1})
cell_format6 = workbook.add_format({"bg_color": 'white',"border":1})
cell_format7 = workbook.add_format({"bg_color": '#FFBF00',"border":1})
cell_format8 = workbook.add_format({"bg_color": '#828282',"border":1})
cell_format9 = workbook.add_format({"bg_color": '#816B56',"border":1})
cell_format10 = workbook.add_format({"bg_color": '#333333',"border":1})
cell_format11 = workbook.add_format({"bg_color": '#00FFBF',"border":1})
cell_format12 = workbook.add_format({"bg_color": '#FFAAD4',"border":1})

for b in range(0,boiteLen):
       w = workbook.add_worksheet(boiteCode[b])
       w.write('A7', 'Etiquette : ' + boiteCode[b], border)
       w.write('O3', date, bold)
       w.write('N3', 'Date de modification : ', bold)
       w.write('A11', 'Entrée', bold)
       w.write('B11', 'Capacité', bold)
       w.write('C11', 'N°         ', bold)
       w.write('D11', 'N° Tube', bold)
       w.write('E11', 'N° Fibre', bold)
       w.write('F11', 'Cassette', bold)
       w.write('G11', 'Etat fibre', bold)
       w.write('H11', 'N° Fibre', bold)
       w.write('I11', 'N° Tube', bold)
       w.write('J11', 'N°       ', bold)
       w.write('K11', 'Capacité', bold)
       w.write('L11', '', bold)
       w.write('M11', 'Sortie', bold)
       w.write('N11', 'Statut', bold)
       w.write('O11', 'Client', bold)
       w.write('A9', 'Reference : ' + boiteReference[b], border)
       w.write('A1', '<- RETOUR:', back)

workbook.close()