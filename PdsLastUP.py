from dbfread import DBF
import xlsxwriter
import datetime

# date configuration
now = datetime.datetime.now()
date = now.strftime("%d/%m/%Y")
# ################## load the both file boite and cable in DBF format ###################################
cableTable = DBF('pdsInput/21_011_076_CABLE_OPTIQUE_B.dbf', load=True, encoding='iso-8859-1')
boiteTable = DBF('pdsInput/21_011_076_BOITE_OPTIQUE_B_AI.dbf', load=True, encoding='iso-8859-1')
zaPboDbl = DBF('pdsInput/zpbodbl.dbf', load=True, encoding='iso-8859-1')
# ################### declare the excel pds file ###########################################################
workbook = xlsxwriter.Workbook('fileGenerated/pds.xlsx')
# ############### define the character and style of cell inside excel ################"
bold = workbook.add_format({'bold': True, "border": 1})
bold1 = workbook.add_format({'bold': True})
border = workbook.add_format({"border": 1})
header = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#037d50'})
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
             cell_format8, cell_format9, cell_format10, cell_format11, cell_format12, border]
# charge the name of all filed in tables
filedCableNam = cableTable.field_names
filedBoiteNam = boiteTable.field_names
boiteLen = len(boiteTable)
cableLen = len(cableTable)
zapLen = len(zaPboDbl)

# #######################declare the table that i need te full#############################################
# FROM THE BOITE OPTIQUE
boiteCode = []  # name of the boite
boiteCable = []  # AMOUNT CABLE
boiteCableState = []  # INTERCO
boiteReference = []  # REFERENCE OF THE BOITE
nbf = []  # NBFUTILE
for j in range(0, boiteLen):
    boiteCode.append(boiteTable.records[j]['NOM'])
    boiteCable.append(boiteTable.records[j]['AMONT'])
    boiteCableState.append(boiteTable.records[j]['INTERCO'])
    boiteReference.append(boiteTable.records[j]['REFERENCE'])
    nbf.append(boiteTable.records[j]['NBFUTILE'])
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
# FROM THE JOIN ZAPBO AND DBL
boiteName = []
nbPrise = []
tECHNO = []
typeBat = []
for k in range(0, zapLen):
    boiteName.append(zaPboDbl.records[k]['NOM'])
    nbPrise.append(zaPboDbl.records[k]['NB_PRISE'])
    tECHNO.append(zaPboDbl.records[k]['TECHNO'])
    typeBat.append(zaPboDbl.records[k]['TYPE_BAT'])


# ########################## functions #################################
def stringCassette(x: str):
    if x.isdigit():
        j = 0
        if int(x) % 12 == 0:
            x = 12
        else:
            x = int(x) % 12

        for i in range(0, 13):
            if i == x:
                x = i
                j = 1
        if j == 1:
            return colorList[x - 1]
        else:
            return colorList[12]
    return colorList[12]


nbmrEpes = 0


def getNumbrFu(boite, nbmrEpes):
    comingBoiteList = []
    for org, extr in zip(cableOrigin, cableExtremity):
        if boite == org:
            comingBoiteList.append(extr)
    y = len(comingBoiteList)
    print(comingBoiteList)
    if y == 0:
        nbmrEpes += nbf[boiteCode.index(boite)]
        print(nbmrEpes)

        print('this', boite)

        return nbmrEpes
    else:
        nbmrEpes += nbf[boiteCode.index(boite)]
        for b in comingBoiteList:
            nbmrEpes = getNumbrFu(b, nbmrEpes)
        return nbmrEpes


def aroundToThree(x: int):
    y = x % 3
    if y != 0:
        k = x + 3 - y
        return k
    else:
        return x


def checkFtt(boit):
    fuFttE = 0
    for b, n, t, y in zip(boiteName, nbPrise, tECHNO, typeBat):
        if boit == b and t == 'FTTE':
            if y == 'PYLONE' or y.startswith('CHT'):
                fuFttE += n * 4
                return aroundToThree(fuFttE)
            else:
                fuFttE += n * 2
                return aroundToThree(fuFttE)
    else:
        return 0


def getSroBoite():
    sroBoite = []
    for o, e in zip(cableOrigin, cableExtremity):
        if o.startswith('SRO'):
            sroBoite.append(e)

    return sroBoite


SROboite = getSroBoite()
print(SROboite)
print('#' * 15)
boite = 'PEC-21-011-076-2000'
print(getNumbrFu(boite, nbmrEpes))
print(checkFtt('PBO-21-011-076-3000'))
