from dbfread import DBF
import xlsxwriter
import datetime

# date configuration
now = datetime.datetime.now()
date = now.strftime("%m/%Y")
SRO = ""
# ################## load the both file boite and cable in DBF format ###################################
cableTable = DBF('etiqueteInputs/21_011_071_CABLE_OPTIQUE_B.dbf', load=True, encoding='iso-8859-1')
boiteTable = DBF('etiqueteInputs/21_011_071_BOITE_OPTIQUE_B.dbf', load=True, encoding='iso-8859-1')
pointTechTable = DBF('etiqueteInputs/21_011_071_POINT_TECHNIQUE_B.dbf', load=True, encoding='iso-8859-1')
supportTable = DBF('etiqueteInputs/21_011_071_SUPPORT_B.dbf', load=True, encoding='iso-8859-1')
filedCableNam = cableTable.field_names
filedBoiteNam = boiteTable.field_names
boiteLen = len(boiteTable)
cableLen = len(cableTable)
pointlen = len(pointTechTable)
supplen = len(supportTable)
# #######################declare the table that i need te full#############################################

# FROM THE BOITE OPTIQUE
boiteCode = []  # name of the boite
boiteAmount = [] # name of the
boiteIdParent = []  # AMOUNT CABLE
for j in range(0, boiteLen):
    boiteCode.append(boiteTable.records[j]['CODE'])
    boiteIdParent.append(boiteTable.records[j]['ID_PARENT'])
    boiteAmount.append(boiteTable.records[j]['AMOUNT'])
# FROM THE CABLE OPTIQUE
cableName = []  # NAME OF THE CABLE
cableOrigin = []  # WHERE THEY COME FROM
cableExtremity = []  # WHERE HE GO IN
cableCapacity = []  # CAPACITY OF THE CABLE
for i in range(0, cableLen):
    cableName.append(cableTable.records[i]['NOM'])
    cableOrigin.append(cableTable.records[i]['ORIGINE'])
    cableExtremity.append(cableTable.records[i]['EXTREMITE'])
    cableCapacity.append(cableTable.records[i]['CAPACITE'])
# FROM THE SUPPORT
suppAmount = []
suppAval = []
suppExi = []
for s in range(0, supplen):
    suppAmount.append(supportTable.records[s]['AMONT'])
    suppAval.append(supportTable.records[s]['AVAL'])
    suppExi.append(supportTable.records[s]['AVAL'])
# FROM THE TECHNIC POINT
pointNom = []
pointCode = []
pointFonc = []
pointStruc = []
pointPrp = []
for p in range(0, pointlen):
    pointNom.append(pointTechTable.records[p]['NOM'])
    pointCode.append(pointTechTable.records[p]['CODE'])
    pointFonc.append(pointTechTable.records[p]['TYPE_FONC'])
    pointStruc.append(pointTechTable.records[p]['TYPE_STRUC'])
    pointPrp.append(pointTechTable.records[p]['PROPRIETAI'])

# #######################################################################################
workbook = xlsxwriter.Workbook('shapeControleResulte'+SRO+'.xlsx')
erourSheet = workbook.add_worksheet("allEroorResulte")

bold = workbook.add_format({'bold': True, "border": 1})
header = workbook.add_format({'bold': True, 'border': 1, 'bg_color': 'red'})
header1 = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#E17070'})
header2 = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#F69494'})