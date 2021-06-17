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
workbook = xlsxwriter.Workbook('PDS/pds-21_011_076.xlsx')
# ############### define the character and style of cell inside excel ################"
bold = workbook.add_format({'bold': True, "border": 1})
bold1 = workbook.add_format({'bold': True})
border = workbook.add_format({"border": 1})
back = workbook.add_format({"bg_color": '#CD5C5C', "border": 1})
header = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#C4E5F7'})
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
boiteFunction = []
nbf = []  # NBFUTILE
for j in range(0, boiteLen):
    boiteCode.append(boiteTable.records[j]['NOM'])
    boiteCable.append(boiteTable.records[j]['AMONT'])
    boiteCableState.append(boiteTable.records[j]['INTERCO'])
    boiteReference.append(boiteTable.records[j]['REFERENCE'])
    boiteFunction.append(boiteTable.records[j]['FONCTION'])
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


def getSroBoite():
    sroBoite = []
    for o, e in zip(cableOrigin, cableExtremity):
        if o.startswith('SRO'):
            sroBoite.append(e)

    return sroBoite


def getListComingBoite(pbo):
    comingList = []
    for org, extr in zip(cableOrigin, cableExtremity):
        if pbo == org:
            comingList.append(extr)
    return comingList


def getboiteOrigine(boite):
    i = boiteCode.index(boite)
    cable = boiteCable[i]
    origin = cableOrigin[cableName.index(cable)]
    return origin


def getNumbrFu(boite, nbmrEpes):
    comingBoiteList = []
    indexB = boiteCode.index(boite)
    capacity = cableCapacity[cableName.index(boiteCable[indexB])]
    fonc = str(boiteFunction[indexB])
    if fonc == 'PEC':
        for org, extr, cap in zip(cableOrigin, cableExtremity, cableCapacity):
            if boite == org:
                if capacity != cap:
                    comingBoiteList.append(extr)
                else:
                    continue
    else:
        for org, extr in zip(cableOrigin, cableExtremity):
            if boite == org:
                comingBoiteList.append(extr)

    y = len(comingBoiteList)
    print(comingBoiteList)
    if y == 0:
        f = nbf[boiteCode.index(boite)]
        if f is None:
            f = 0
        nbmrEpes += f
        print(nbmrEpes)

        print('this', boite)

        return nbmrEpes
    else:
        f = nbf[boiteCode.index(boite)]
        if f is None:
            f = 0
        nbmrEpes += f
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
        if boit == b:
            if t == 'FTTE':
                if y == 'PYLONE' or y.startswith('CHT'):
                    fuFttE += n * 4
                    return aroundToThree(fuFttE)
                else:
                    fuFttE += n * 2
                    return aroundToThree(fuFttE)
            elif t == 'FTTH' and y == 'BATIMENT PUBLIC':
                fuFttE += n * 2
                return aroundToThree(fuFttE)
    else:
        return 0


def checkGlobalFtt(bo):
    listBoit = getListComingBoite(bo)
    x = 0
    if len(listBoit) == 0:
        return checkFtt(bo)
    else:
        x += checkFtt(bo)
        for pbo in listBoit:
            x += checkGlobalFtt(pbo)
        return x


sheet = xlsxwriter.worksheet.Worksheet


def baseSheet(boite, w: sheet):
    # INFORMATION ABOUT BOITE
    # boite name
    w.write('Q1', 'Etiquette : ', header)
    w.write('R1', boite, bold)
    # boite Ref
    w.write('Q2', 'Reference : ', header)
    w.write('R2', boiteReference[boiteCode.index(boite)], bold)
    # date Now
    w.write('Q3', 'Date de modification : ', header)
    w.write('R3', date, bold)
    # boite Origine
    w.write('Q5', 'RETURN : ', back)
    orgin = getboiteOrigine(boite)
    w.write('Q6', orgin, bold)
    # boite Next boite coming
    w.write('R5', 'NEXT : ', back)
    BoiteNext = getListComingBoite(boite)
    R = 6
    for l in BoiteNext:
        w.write('R' + str(R), l, bold)
        R += 1

    # INFORMATION OF THE HEADER
    w.write('A1', 'Entrée', header)
    w.write('B1', 'Capacité', header)
    w.write('C1', 'N°         ', header)
    w.write('D1', 'N° Tube', header)
    w.write('E1', 'N° Fibre', header)
    w.write('F1', 'Cassette', header)
    w.write('G1', 'Etat fibre', header)
    w.write('H1', 'N° Fibre', header)
    w.write('I1', 'N° Tube', header)
    w.write('J1', 'N°       ', header)
    w.write('K1', 'Capacité', header)
    w.write('L1', '', header)
    w.write('M1', 'Sortie', header)
    w.write('N1', 'Statut', header)
    w.write('O1', 'Client', header)


def getCapacity(cable):
    i = cableName.index(cable)
    capacity = cableCapacity[i]
    return capacity


def cableBaseInfo(w: sheet, cable, capacity, T=1, ):
    for i in range(0, capacity):

        w.write(i + 1, 0, cable, border)
        w.write(i + 1, 1, capacity, border)
        num = (i % 12) + 1
        w.write(i + 1, 2, num, border)
        w.write(i + 1, 3, T, colorList[T - 1])
        if num % 12 == 0:
            if T == 12:
                T = 1
            else:
                T += 1
        w.write(i + 1, 4, num, colorList[num - 1])


SROboite = getSroBoite()
print(SROboite)
# print('#' * 15)
print(getNumbrFu('PBO-21-011-076-2007', 0))
print('#' * 15)
print(checkGlobalFtt('PEC-21-011-076-2000'))

# ########################################## start fill In the pds ##############################

for b in range(0, boiteLen):
    # ################## constant work with ####################
    N = 1
    T = 1
    Len = 0
    F = 0
    stockN = 0
    fuNumber = 0
    ftte = 0
    nbrEpesSansFTTE = 0
    # ##########################################################
    w = workbook.add_worksheet(boiteCode[b])
    boite = boiteCode[b]
    func = boiteFunction[b]
    baseSheet(boite, w)
    cable = boiteCable[b]
    capacity = getCapacity(cable)
    if func == 'PEC':
        stockN = 0
        fuNumber = getNumbrFu(boite, 0)
        ftte = checkGlobalFtt(boite)
        nbrEpesSansFTTE = fuNumber - ftte
    else:
        stockN = nbf[b]
        ffuNumber = getNumbrFu(boite, 0)
        nbrEpesSansFTTE = fuNumber - stockN
    cableBaseInfo(w,cable,capacity,T)

workbook.close()
