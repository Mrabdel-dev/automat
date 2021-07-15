import operator
import string

from dbfread import DBF, FieldParser
import xlsxwriter
import datetime


class MyFieldParser(FieldParser):
    def parseN(self, field, data):
        data = data.strip().strip(
            b'*\x00')  # Had to strip out the other characters first before \x00, as per super function specs.
        return super(MyFieldParser, self).parseN(field, data)

    def parseD(self, field, data):
        data = data.strip(b'\x00')
        return super(MyFieldParser, self).parseD(field, data)


# date configuration
now = datetime.datetime.now()
date = now.strftime("%d/%m/%Y")
# ################## load the both file boite and cable in DBF format ###################################
cableTable = DBF('pdsInput/21_011_075_CABLE_OPTIQUE_B.dbf', load=True, encoding='iso-8859-1')
boiteTable = DBF('pdsInput/21_011_075_BOITE_OPTIQUE_B.dbf', load=True, encoding='iso-8859-1')
zaPboDbl = DBF('pdsInput/ZAPBO-DBL.dbf', load=True, encoding='iso-8859-1')
casseteTable = DBF('pdsInput/cassete_file.dbf', load=True, encoding='iso-8859-1')
# ################### declare the excel pds file ###########################################################
workbook = xlsxwriter.Workbook('PDS/SRO-21_011_075-Pds.xlsx')
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
cassLen = len(casseteTable)

# #######################declare the table that i need te full#############################################
# FROM THE BOITE OPTIQUE
boiteCode = []  # name of the boite
boiteCable = []  # AMOUNT CABLE
boiteCableState = []  # INTERCO
boiteReference = []  # REFERENCE OF THE BOITE
boiteFunction = []  # boite Func { PEC OR PEC-PBO OR PBO)
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
statut = []
for k in range(0, zapLen):
    boiteName.append(zaPboDbl.records[k]['NOM'])
    nbPrise.append(zaPboDbl.records[k]['NB_PRISE'])
    tECHNO.append(zaPboDbl.records[k]['TECHNO'])
    typeBat.append(zaPboDbl.records[k]['TYPE_BAT'])
    statut.append(zaPboDbl.records[k]['STATUT'])

# from the cassete file
reference = []  # reference of the boite
nbrCassete = []  # nbr cassete dans la boite
tailleCassete = []  # nbr de fibre dans chaque cassete
for c in range(0, cassLen):
    reference.append(casseteTable.records[c]['REF'])
    nbrCassete.append(casseteTable.records[c]['NBR_CASS'])
    tailleCassete.append(casseteTable.records[c]['TAILLE'])

sheet = xlsxwriter.worksheet.Worksheet

# ########################## functions #################################
nbmrEpes = 0


# color of tube or fibre
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


# get all Sro boite
def getSroBoite():
    sroBoite = []
    for o, e in zip(cableOrigin, cableExtremity):
        if o.startswith('SRO'):
            sroBoite.append(e)

    return sroBoite


# function return all next coming boite
def getListComingBoite(pbo):
    comingList = []
    for org, extr in zip(cableOrigin, cableExtremity):
        if pbo == org:
            comingList.append(extr)
    dectBoit = {}
    for b in comingList:
        nbfu = getfuNum(b, 0)
        dectBoit.update({b: nbfu})
    comingL = {k: v for k, v in sorted(dectBoit.items(), key=lambda v: v[1], reverse=True)}
    comingList = list(comingL.keys())
    return comingList


# function return all the next withe boite tha PIC with some capacity
def getListComingBoitePEC(pbo):
    comingList = []
    ind = boiteCode.index(pbo)
    cable = boiteCable[ind]
    capaci = getCapacity(cable)
    for org, extr in zip(cableOrigin, cableExtremity):
        if pbo == org:
            index = boiteCode.index(extr)
            cable1 = boiteCable[index]
            capcity2 = getCapacity(cable1)
            if capaci != capcity2:
                comingList.append(extr)
    dectBoit = {}
    for b in comingList:
        nbfu = getfuNum(b, 0)
        dectBoit.update({b: nbfu})
        print(dectBoit)
    comingL = {k: v for k, v in sorted(dectBoit.items(), key=lambda v: v[1], reverse=True)}
    comingList = list(comingL.keys())
    return comingList


# functio return the boite origine of a specific boite
def getboiteOrigine(boite):
    i = boiteCode.index(boite)
    cable = boiteCable[i]
    origin = cableOrigin[cableName.index(cable)]
    return origin


# get the rsulte of fu of next boite
def getNumbrFu(boite, nbmrEpes):
    comingBoiteList = []
    indexB = boiteCode.index(boite)
    fonc = str(boiteFunction[indexB])
    comingBoiteList = getListComingBoitePEC(boite)
    # if fonc == 'PBO':
    #     comingBoiteList = getListComingBoite(boite)

    y = len(comingBoiteList)

    if y == 0:
        f = nbf[boiteCode.index(boite)]
        if f is None:
            f = 0
        nbmrEpes += f
        return nbmrEpes
    else:
        f = nbf[boiteCode.index(boite)]
        if f is None:
            f = 0
        nbmrEpes += f
        for b in comingBoiteList:
            nbmrEpes = getfuNum(b, nbmrEpes)
        return nbmrEpes


def getfuNum(boite, nbmrEpes):
    comingBoiteList = getListComingBoite(boite)
    y = len(comingBoiteList)

    if y == 0:
        f = nbf[boiteCode.index(boite)]
        if f is None:
            f = 0
        nbmrEpes += f
        return nbmrEpes
    else:
        f = nbf[boiteCode.index(boite)]
        if f is None:
            f = 0
        nbmrEpes += f
        for b in comingBoiteList:
            nbmrEpes = getfuNum(b, nbmrEpes)
        return nbmrEpes


# math function to major a number to a specific num
def aroundTo(x: int, num):
    if x == 0:
        x = 1
    y = x % num
    if y != 0:
        k = x + num - y
        return k
    else:
        return x


# get fu ftte of a boite
def checkFtt(boit):
    fuFttE = 0
    for b, n, t, y, s in zip(boiteName, nbPrise, tECHNO, typeBat, statut):
        if boit == b:
            if t == 'FTTE' and s != 'ABANDONNE':
                if y == 'PYLONE' or y.startswith('CHT'):
                    fuFttE += n * 4

                else:
                    fuFttE += n * 2

            elif t == 'FTTH' and (y == 'BATIMENT PUBLIC' or y == 'BATIMENT RELIGIEUX'):
                fuFttE += n * 2
    if fuFttE != 0:
        return aroundTo(fuFttE, 3)
    else:
        return 0


# get the resulte ftte of all next coming boite
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


# founction to capcity of cable
def getCapacity(cable):
    i = cableName.index(cable)
    capacity = int(cableCapacity[i])
    return capacity


# function to cable index based on boite name
def getCableIndex(boite):
    index = boiteCode.index(boite)
    cable = boiteCable[index]
    indexc = cableName.index(cable)
    return indexc


# function to cable based on boite
def getCable(boite):
    index = boiteCode.index(boite)
    cable = boiteCable[index]
    return cable


# function to get last boite have some capacity cable
def getLastStartBoite(boite):
    index = getCableIndex(boite)
    capacity = cableCapacity[index]
    orginBoite = cableOrigin[index]
    if orginBoite.startswith('SRO'):
        return boite
    else:
        index2 = getCableIndex(orginBoite)
        capacity2 = cableCapacity[index2]
        if capacity == capacity2:
            try:
                return getLastStartBoite(orginBoite)
            except ValueError:
                return orginBoite
        else:
            return boite


def getAllboitestart(boitestart, boite, listB):
    cab = getCable(boite)
    cap = getCapacity(cab)
    if boitestart == boite:
        listB.append(boite)
        return listB
    boites = getListComingBoite(boitestart)
    listB.append(boitestart)
    for b in boites:
        cab1 = getCable(b)
        cap1 = getCapacity(cab1)
        if cap1 == cap:
            return getAllboitestart(b, boite, listB)


# function return where i should start write to write stocked state
def getStockStartLine(boite):
    cab = getCable(boite)
    cap = getCapacity(cab)
    fuUsed = 0
    b = getLastStartBoite(boite)
    listB = []
    listB = getAllboitestart(b, boite, listB)
    for l in listB:
        fuUsed += getNumbrFu(l, 0)
    fuBoit = getNumbrFu(boite, 0)
    fuUsed -= getPassedFtte(boite, cap)
    lineStart = fuUsed - fuBoit
    return lineStart


# function return where i should start write to write stocked state
def getBoitePassage(boite):
    cable = getCable(boite)
    cap = getCapacity(cable)
    listBoits = getListComingBoite(boite)
    for b in listBoits:
        cab = getCable(b)
        capc = getCapacity(cab)
        if cap == capc:
            return b
    return None


def getPassedFtte(boite, capacity):
    startboite = getLastStartBoite(boite)
    totalFTTE = 0
    test = True
    while test:
        if startboite == boite:
            test = False
        else:
            totalFTTE += checkFtt(startboite)
            listBoite = getListComingBoite(startboite)
            if len(listBoite) > 0:
                for b in listBoite:
                    cab = getCable(b)
                    cap = getCapacity(cab)
                    if cap == capacity:
                        startboite = b
    return totalFTTE


# function return where i should start write ftte
def getFTTElineStart(boite):
    cable = getCable(boite)
    capacity = getCapacity(cable)
    x = checkGlobalFtt(boite) + getPassedFtte(boite, capacity)
    line = capacity - aroundTo(x, 12)
    return line + 1


# function to get boit that have ftte prise
def getFTTEBoites(boite, Pass: bool):
    if Pass:
        listBoit = getListComingBoitePEC(boite)
    else:
        listBoit = getListComingBoite(boite)
    listFFTE = []
    for i in listBoit:
        x = checkGlobalFtt(i)
        if x != 0:
            listFFTE.append(i)
    return listFFTE


# function write the basic header for the sheet
def baseSheet(w: sheet, boite):
    # INFORMATION ABOUT BOITE
    # boite name
    # w.freeze_panes = w['R']
    # w.freeze_panes = w['Q']

    w.freeze_panes(0, 1)
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
    if str(orgin).startswith('SRO'):
        w.write('Q6', orgin, bold)
    else:
        w.write_url('Q6', f"internal:'{orgin}'!R1", string=orgin)
    # boite Next boite coming
    w.write('R5', 'NEXT : ', back)
    BoiteNext = getListComingBoite(boite)
    k = len(BoiteNext)
    if k > 0:
        R = 6
        for l in BoiteNext:
            w.write_url('R' + str(R), f"internal:'{l}'!R1", string=l)
            R += 1
    else:
        l = 'EXTREMITE'
        w.write('R6', l, bold)
    # for i in range(0,15):
    #     w.freeze_panes =w[str(string.ascii_uppercase[i])+"1"]

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


# function to write the basic info of the boite and cable
def cableBaseInfo(w: sheet, cable, capacity, T=1, ):
    for i in range(0, capacity):

        w.write(i + 1, 0, cable, border)
        w.write(i + 1, 1, capacity, border)
        num = (i % 12) + 1
        w.write(i + 1, 3, T, stringCassette(str(T)))
        if num % 12 == 0:
            if T == 24:
                T = 1
            else:
                T += 1
        w.write(i + 1, 4, num, stringCassette(str(num)))
        w.write(i + 1, 5, '', border)


# function to write next cable epesuree on the boite just for specific next boit
def fillInEpess(w: sheet, Lin, i, boite, T, N, k, size, p):
    cable = getCable(boite)
    capacity = getCapacity(cable)
    ftt = checkGlobalFtt(boite)
    funb = getfuNum(boite, 0)
    nbrEps = funb - ftt
    N = int(N)
    for j in range(0, nbrEps):
        if N < 10:
            n = 'CSE-0' + str(N)
        else:
            n = 'CSE-' + str(N)

        w.write(Lin, 2, p, border)
        w.write(Lin, 5, n, border)
        k += 1
        w.write(Lin, 6, 'EPISSUREE', border)
        w.write(Lin, 10, capacity, border)
        num = (i % 12) + 1
        w.write(Lin, 9, num, border)
        w.write(Lin, 8, T, stringCassette(str(T)))
        if num % 12 == 0:
            if T == 24:
                T = 1
            else:
                T += 1
        w.write(Lin, 7, num, stringCassette(str(num)))
        w.write(Lin, 12, cable, border)
        w.write(Lin, 11, '', border)
        w.write(Lin, 13, 'EPISSUREE', border)
        w.write(Lin, 14, '', border)
        p += 1
        if k > size:
            k = 1
            p = 1
            N += 1
        i += 1
        Lin += 1
    return Lin, k, N, p


# function to write  all next cable epesuree on the boite
def fillInAllCableEpess(w: sheet, nextBoite, boite, Lin):
    index = getcassteIndex(boite)
    size = tailleCassete[index]
    ftte = checkGlobalFtt(boite)
    N = aroundTo(ftte, size) / size
    k = 1
    p = 1
    for b in nextBoite:
        x, k, N, p = fillInEpess(w, Lin, 0, b, 1, N, k, size, p)
        Lin = x

    return Lin


# function  to write the ftte nex cable
def ftteFillIn(w, Listboites, boite, startLin, T):
    index = getcassteIndex(boite)
    size = tailleCassete[index]
    k = 1
    N = 1
    p = 1
    for b in Listboites:
        i = 0
        ftteN = checkGlobalFtt(b)
        cable = getCable(b)
        capacity = getCapacity(cable)
        T = tubeRound(capacity - ftteN)
        for j in range(0, ftteN):
            if startLin % 12 == 0:
                startLin = startLin - 2 * 12
            w.write(startLin, 2, p, border)
            if N < 10:
                n = 'CSE-0' + str(N)
            else:
                n = 'CSE-' + str(N)
            w.write(startLin, 5, n, border)
            k += 1
            w.write(startLin, 6, 'EPISSUREE', border)
            w.write(startLin, 10, capacity, border)
            num = (i % 12) + 1
            w.write(startLin, 9, num, border)
            w.write(startLin, 8, T, stringCassette(str(T)))
            if num % 12 == 0:
                if T == 24:
                    T = 1
                else:
                    T += 1
            w.write(startLin, 7, num, stringCassette(str(num)))
            w.write(startLin, 12, cable, border)
            w.write(startLin, 11, '', border)
            w.write(startLin, 13, 'EPISSUREE', border)
            w.write(startLin, 14, '', border)
            i += 1
            p += 1
            if k > size:
                k = 1
                p = 1
                N += 1
            startLin += 1

    return startLin


def tubeRound(num):
    T = 1
    for i in range(0, num):
        x = (i % 12) + 1
        if x % 12 == 0:
            if T == 24:
                T = 1
            else:
                T += 1
    return T


def fillfPassedfttePassage(w, boite, nbrPassFTTE, T):
    cable = getCable(boite)
    cap = getCapacity(cable)
    n = 1
    startLine = (cap - 12 * n) + 1
    i = startLine - 1
    index = getcassteIndex(boite)
    size = tailleCassete[index]
    k = 1
    N = 1
    p = 1
    for f in range(0, nbrPassFTTE):
        if N < 10:
            e = 'CSE-0' + str(N)
        else:
            e = 'CSE-' + str(N)
        w.write(startLine, 5, e, border)
        w.write(startLine, 6, 'EN PASSAGE', border)
        num = (i % 12) + 1
        w.write(startLine, 8, T, stringCassette(str(T)))
        if num % 12 == 0:
            if T == 24:
                T = 1
            else:
                T += 1
        w.write(startLine, 7, num, stringCassette(str(num)))
        w.write(startLine, 9, num, border)
        w.write(startLine, 10, cap, border)
        w.write(startLine, 11, '', border)
        w.write(startLine, 12, cable, border)
        w.write(startLine, 14, '', border)
        w.write(startLine, 13, 'EN PASSAGE', border)
        startLine += 1
        i += 1
        if k > size:
            k = 1
            N += 1
        if (startLine - 1) % 12 == 0:
            n += 1
            startLine = (cap - 12 * n) + 1
    return startLine


# function  to write the passage  next cable
def fillPecPassage(w, boite, startLine, endLine, i, T, p):
    cable = getCable(boite)
    cap = getCapacity(cable)
    for k in range(startLine, endLine):
        w.write(startLine, 2, p, border)
        w.write(startLine, 5, 'FOND DE BOITE', border)
        w.write(startLine, 6, 'EN PASSAGE', border)
        num = (i % 12) + 1
        w.write(startLine, 8, T, stringCassette(str(T)))
        if num % 12 == 0:
            if T == 24:
                T = 1
            else:
                T += 1
        w.write(startLine, 7, num, stringCassette(str(num)))
        w.write(startLine, 9, num, border)
        w.write(startLine, 10, cap, border)
        w.write(startLine, 11, '', border)
        w.write(startLine, 12, cable, border)
        w.write(startLine, 14, '', border)
        w.write(startLine, 13, 'EN PASSAGE', border)
        startLine += 1
        i += 1
        p += 1
    return p


def PboFillFTTeStocker(w: sheet, boite, stokker, Lin, T=1):
    i = Lin - 1
    index = getcassteIndex(boite)
    N = nbrCassete[index]
    fu = getNumbrFu(boite, 0)
    ftte = checkGlobalFtt(boite)
    y = fu - ftte
    if y > 0:
        N = N - int(aroundTo(y, N) / N)
    else:
        N = nbrCassete[index]

    size = tailleCassete[index]
    k = 1
    p = 1
    for s in range(0, stokker):
        if Lin % 12 == 0:
            Lin = Lin - 2 * 12
        if N < 10:
            n = 'CSE-0' + str(N)
        else:
            n = 'CSE-' + str(N)
        w.write(Lin, 2, p, border)
        w.write(Lin, 5, n, border)
        k += 1
        w.write(Lin, 6, 'STOCKEE', border)
        w.write(Lin, 10, '', border)
        num = (i % 12) + 1
        w.write(Lin, 9, '', border)
        w.write(Lin, 8, '', border)
        if num % 12 == 0:
            if T == 12:
                T = 1
            else:
                T += 1
        w.write(Lin, 7, '', border)
        w.write(Lin, 11, '', border)
        w.write(Lin, 12, '', border)
        w.write(Lin, 14, '', border)
        w.write(Lin, 13, 'STOCKEE', border)
        p += 1
        if k > size:
            N = N - 1
            p = 1
            k = 1
        Lin += 1
        i += 1
    return Lin


# function to write stoker state
def PboFillStokker(w: sheet, boite, stokker, Lin, T=1):
    i = Lin - 1
    index = getcassteIndex(boite)
    N = nbrCassete[index]
    size = tailleCassete[index]
    k = 1
    p = 1
    for s in range(0, stokker):
        if N < 10:
            n = 'CSE-0' + str(N)
        else:
            n = 'CSE-' + str(N)
        w.write(Lin, 2, p, border)
        w.write(Lin, 5, n, border)
        k += 1
        w.write(Lin, 6, 'STOCKEE', border)
        w.write(Lin, 10, '', border)
        num = (i % 12) + 1
        w.write(Lin, 9, '', border)
        w.write(Lin, 8, '', border)
        if num % 12 == 0:
            if T == 12:
                T = 1
            else:
                T += 1
        w.write(Lin, 7, '', border)
        w.write(Lin, 11, '', border)
        w.write(Lin, 12, '', border)
        w.write(Lin, 14, '', border)
        w.write(Lin, 13, 'STOCKEE', border)
        p += 1
        if k > size:
            N = N - 1
            k = 1
            p = 1
        Lin += 1
        i += 1
    return Lin


# function to write epssure state for PEC-PBO
def PboFillEpes(w: sheet, boites, boite, Lin, i, T=1):
    index = getcassteIndex(boite)
    size = tailleCassete[index]
    nbrCas = nbrCassete[index]
    ftte = checkGlobalFtt(boite)
    N = aroundTo(ftte, size) / size
    k = 1
    p = 1
    for s in boites:
        x, k, N, p = fillInEpess(w, Lin, i, s, 1, N, k, size, p)
        Lin = x
    return Lin


# function to write all passage state next cable
def passageFillIn(w: sheet, boit, startLine, T=1):
    boitlist = getListComingBoite(boit)
    for b in boitlist:
        cable = getCable(b)
        capacity = getCable(cable)
        nmbrfu = getNumbrFu(b, 0)
        i = 0
        for k in range(0, nmbrfu):
            w.write(startLine, 5, 'FOND DE BOITE', border)
            w.write(startLine, 6, 'EN PASSAGE', border)
            w.write(startLine, 10, capacity, border)
            num = (i % 12) + 1
            w.write(startLine, 9, num, border)
            w.write(startLine, 8, T, stringCassette(str(T)))
            if num % 12 == 0:
                if T == 12:
                    T = 1
                else:
                    T += 1
            w.write(startLine, 7, num, stringCassette(str(num)))
            w.write(startLine, 12, cable, border)
            w.write(startLine, 13, 'EN PASSAGE', border)
            startLine += 1
            i += 1
    return startLine


def librePassFTTEFill(w: sheet, boit, fttePass, T=1):
    cable = getCable(boite)
    cap = getCapacity(cable)
    n = 1
    startLine = (cap - 12 * n) + 1
    i = startLine - 1
    for k in range(0, fttePass):
        w.write(startLine, 5, 'FOND DE BOITE', border)
        w.write(startLine, 6, 'LIBRE', border)
        w.write(startLine, 8, '', border)
        w.write(startLine, 7, '', border)
        w.write(startLine, 9, '', border)
        w.write(startLine, 10, '', border)
        w.write(startLine, 11, '', border)
        w.write(startLine, 12, '', border)
        w.write(startLine, 14, '', border)
        w.write(startLine, 13, 'LIBRE', border)
        startLine += 1
        i += 1
        if (startLine - 1) % 12 == 0:
            n += 1
            startLine = (cap - 12 * n) + 1
    return startLine


# function to write the libre state for next cable
def libreFillIn(w: sheet, boit, startLine, endLine, p: int):
    i = 1
    for k in range(startLine, endLine):
        w.write(startLine, 2, p, border)
        w.write(startLine, 5, 'FOND DE BOITE', border)
        w.write(startLine, 6, 'LIBRE', border)
        w.write(startLine, 8, '', border)
        w.write(startLine, 7, '', border)
        w.write(startLine, 9, '', border)
        w.write(startLine, 10, '', border)
        w.write(startLine, 11, '', border)
        w.write(startLine, 12, '', border)
        w.write(startLine, 14, '', border)
        w.write(startLine, 13, 'LIBRE', border)
        startLine += 1
        i += 1
        p += 1
    return p


def extracableFillIn(w: sheet, cable, cap, extarline, startLine, funm, p):
    i = funm
    T = tubeRound(funm)
    for e in range(0, extarline):
        w.write(startLine, 0, '', border)
        w.write(startLine, 1, '', border)
        w.write(startLine, 2, p, border)
        w.write(startLine, 3, '', border)
        w.write(startLine, 4, '', border)
        w.write(startLine, 5, 'FOND DE BOITE', border)
        w.write(startLine, 6, 'LIBRE', border)
        num = (i % 12) + 1
        w.write(startLine, 7, num, stringCassette(str(num)))
        w.write(startLine, 8, T, stringCassette(str(T)))
        w.write(startLine, 9, num, border)
        if num % 12 == 0:
            if T == 12:
                T = 1
            else:
                T += 1
        w.write(startLine, 10, cap, border)
        w.write(startLine, 11, '', border)
        w.write(startLine, 12, cable, border)
        w.write(startLine, 13, 'LIBRE', border)
        w.write(startLine, 14, '', border)
        startLine += 1
        i += 1
        p += 1
    return startLine, p


# function to write all extract libre cable need for sorted cable
def extracablePECPBOFillIn(w: sheet, boites, boite, startLine, p):
    y = getBoitePassage(boite)
    index1 = boiteCode.index(boite)
    func = boiteFunction[index1]
    k = getFTTEBoites(boite, True)
    fuNumbr = nbf[index1]
    fuNumbr1 = getNumbrFu(boite, 0) - 1
    test = False
    for b in boites:
        for l in k:
            if b == l:
                test = True
            else:
                test = False

        if not test:
            if y != b:
                fuN = getfuNum(b, 0)
                cable = getCable(b)
                cap = getCapacity(cable)
                extraN = cap - fuN
                startLine, p = extracableFillIn(w, cable, cap, extraN, startLine, fuN, p)

            else:
                cable = getCable(b)
                cap = getCapacity(cable)
                fu = getNumbrFu(b, 0)
                total = fuNumbr1
                if func == 'PEC':
                    start = getLastStartBoite(boite)
                    if start == boite:
                        Lin = 1
                    else:
                        Lin = getNumbrFu(getLastStartBoite(start), 0)
                    startLine, p = extracableFillIn(w, cable, cap, fuNumbr1, startLine, Lin, p)
                else:
                    Lin = getStockStartLine(boite)
                    startLine, p = extracableFillIn(w, cable, cap, total, startLine, Lin, p)


        else:
            if y != b:
                cable = getCable(b)
                cap = getCapacity(cable)
                ftte = checkGlobalFtt(b)
                nbfu = getfuNum(b, 0) - ftte
                extraN2 = cap - aroundTo(ftte, 12) - getfuNum(b, 0) + ftte
                startLine, p = extracableFillIn(w, cable, cap, extraN2, startLine, nbfu, p)
                extraN = ftte
                tt = cap - aroundTo(ftte, 12) + ftte
                startLine, p = extracableFillIn(w, cable, cap, extraN, startLine, tt, p)

            else:
                cable = getCable(b)
                cap = getCapacity(cable)
                ftte = checkGlobalFtt(b)
                fu = getfuNum(b, 0)
                total = fuNumbr1
                if func == 'PEC':
                    start = getLastStartBoite(boite)
                    if start == boite:
                        Lin = 1
                    else:
                        Lin = getNumbrFu(getLastStartBoite(start), 0)
                    startLine, p = extracableFillIn(w, cable, cap, fuNumbr1, startLine, Lin, p)
                else:
                    Lin = getStockStartLine(boite)
                    startLine, p = extracableFillIn(w, cable, cap, total, startLine, Lin, p)
                    # lin = Lin + getNumbrFu(getLastStartBoite(b), 0)-ftte
                    # startLine = extracableFillIn(w, cable, cap, fuNumbr1-ftte, startLine, Lin)


listCasseteNotfound = []


def getcassteIndex(boite):
    index = boiteCode.index(boite)
    ref = boiteReference[index]
    try:
        indexCass = reference.index(ref)
        return indexCass
    except ValueError:
        listCasseteNotfound.append(ref)
        indexCass = 0
        return indexCass


def cassteFillIn(w: sheet, boite, function):
    index = boiteCode.index(boite)
    ref = boiteReference[index]
    ftte = checkGlobalFtt(boite)
    cassIndex = getcassteIndex(boite)
    if function == 'PEC':
        pass
    elif function == 'PEC_PBO':
        pass
    else:
        pass


def passageCasseteFillIn(w: sheet, boites, line, size, cass):
    i = 0

    for b in boites:
        ftte = checkGlobalFtt(b)
        fu = getNumbrFu(b, 0) - ftte
        for k in range(0, fu):
            T = 'CSE-' + str(cass)
            w.write(line, 5, T, border)
            i += 1
            line += 1
            if i > size:
                cass += 1


# <--################### PEC function #############################-->
def boitePecFillIn(w: sheet, cable, boite, capacity, T):
    baseSheet(w, boite)
    p = 1
    fuNumber = getNumbrFu(boite, 0)
    fttepass = getPassedFtte(boite, capacity)
    ftte = checkGlobalFtt(boite)
    fttetotal = ftte + fttepass
    ftteLineD = getFTTElineStart(boite)
    Test = (fuNumber - checkGlobalFtt(boite)) + aroundTo(fttetotal, 12)
    if Test > capacity:
        print(boite, "Erouuuuuuuuur cable capacity not enough")
    cableBaseInfo(w, cable, capacity, T)
    nextBoits = getListComingBoitePEC(boite)
    boites = getListComingBoite(boite)
    start = getLastStartBoite(boite)
    print(start)
    if start == boite:
        Lin = 1
    else:
        Lin = (getNumbrFu(getLastStartBoite(start), 0) - (checkGlobalFtt(start) - checkGlobalFtt(boite))) + 1
        print(boite, Lin)
    x = getBoitePassage(boite)
    if x is not None:
        if ftte != 0:
            p = fillPecPassage(w, x, 1, Lin, Lin - 1, tubeRound(Lin), p)
            Lin = fillInAllCableEpess(w, nextBoits, boite, Lin)
            ftteLine = fillfPassedfttePassage(w, boite, fttepass, T)
            endFTTLine = ftteFillIn(w, nextBoits, boite, ftteLine, T)
            end = aroundTo(endFTTLine, 12)
            p = fillPecPassage(w, x, Lin, ftteLineD, Lin - 1, tubeRound(Lin), p)
            p = fillPecPassage(w, x, endFTTLine, end + 1, endFTTLine - 1, tubeRound(endFTTLine), p)
        else:
            p = fillPecPassage(w, x, 1, Lin, Lin - 1, tubeRound(Lin), p)
            Lin = fillInAllCableEpess(w, nextBoits, boite, Lin)
            p = fillPecPassage(w, x, Lin, capacity + 1, Lin - 1, tubeRound(Lin), p)

    else:
        if ftte != 0:
            p = libreFillIn(w, boite, 1, Lin, p)
            Lin = fillInAllCableEpess(w, boites, boite, Lin)
            ftteLine = librePassFTTEFill(w, boite, fttepass, T)
            endFTTLine = ftteFillIn(w, boites, boite, ftteLine, T)
            end = aroundTo(endFTTLine, 12)
            p = libreFillIn(w, boite, Lin, ftteLineD, p)
            p = libreFillIn(w, boite, endFTTLine, end + 1, p)
        else:
            p = libreFillIn(w, boite, 1, Lin, p)
            Lin = fillInAllCableEpess(w, boites, boite, Lin)
            p = libreFillIn(w, boite, Lin, capacity + 1, p)
    extracablePECPBOFillIn(w, boites, boite, capacity + 1, p)


# <--################### PEC-PBO function ##########################-->
def boitePecPboFillIn(w: sheet, cable, boite, capacity, T):
    endftteLine = 0
    ftteLine = 0
    p = 1
    baseSheet(w, boite)
    indexCass = getcassteIndex(boite)
    cableBaseInfo(w, cable, capacity, T)
    linestockstart = getStockStartLine(boite) + 1
    index = boiteCode.index(boite)
    ftt = checkFtt(boite)
    ftte = checkGlobalFtt(boite)
    fttepass = getPassedFtte(boite, capacity)
    ftteLine = getFTTElineStart(boite)
    stoker = nbf[index] - checkFtt(boite)
    Lin = PboFillStokker(w, boite, stoker, linestockstart, 1)
    boites = getListComingBoitePEC(boite)
    boitesF = getListComingBoite(boite)
    x = getBoitePassage(boite)
    if x is not None:
        if ftte != 0:
            Lin = PboFillEpes(w, boites, boite, Lin, 0, 1)
            fttePas = fillfPassedfttePassage(w, boite, fttepass, T)
            endftteLineS = PboFillFTTeStocker(w, boite, ftt, fttePas, T)
            endftteLine = ftteFillIn(w, boites, boite, endftteLineS, 1)
            end = aroundTo(endftteLine, 12)
            p = fillPecPassage(w, x, 1, linestockstart, 0, T, p)
            p = fillPecPassage(w, x, Lin, ftteLine, Lin - 1, tubeRound(Lin), p)
            p = fillPecPassage(w, x, endftteLine, end + 1, endftteLine - 1, tubeRound(endftteLine), p)
        else:
            Lin = PboFillEpes(w, boites, boite, Lin, 0, 1)
            p = fillPecPassage(w, x, 1, linestockstart, 0, T, p)
            p = fillPecPassage(w, x, Lin, capacity + 1, Lin - 1, tubeRound(Lin), p)
    else:
        if ftte != 0:
            Lin = PboFillEpes(w, boitesF, boite, Lin, 0, 1)
            fttePas = librePassFTTEFill(w, boite, fttepass, T)
            endftteLineS = PboFillFTTeStocker(w, boite, ftt, fttePas, T)
            endftteLine = ftteFillIn(w, boitesF, boite, endftteLineS, 1)
            end = aroundTo(endftteLine, 12)
            p = libreFillIn(w, boite, 1, linestockstart, p)
            p = libreFillIn(w, boite, Lin, ftteLine, p)
            p = libreFillIn(w, boite, endftteLine, end + 1, p)
        else:
            Lin = PboFillEpes(w, boitesF, boite, Lin, 0, 1)
            p = libreFillIn(w, boite, 1, linestockstart, p)
            p = libreFillIn(w, boite, Lin, capacity + 1, p)

    extracablePECPBOFillIn(w, boitesF, boite, capacity + 1, p)


# <--################### PBO function #############################-->
def boitePboFillIn(w: sheet, cable, boite, capacity, T):
    baseSheet(w, boite)
    p = 1
    indexCass = getcassteIndex(boite)
    cableBaseInfo(w, cable, capacity, T)
    linestockstart = getStockStartLine(boite) + 1
    index = boiteCode.index(boite)
    ftte = checkGlobalFtt(boite)
    ftt = checkFtt(boite)
    fttepass = getPassedFtte(boite, capacity)
    stoker = nbf[index] - checkFtt(boite)
    Lin = PboFillStokker(w, boite, stoker, linestockstart, 1)
    ftteLine = getFTTElineStart(boite)
    boites = getListComingBoite(boite)
    boitsF = getListComingBoitePEC(boite)
    x = getBoitePassage(boite)
    if x is not None:
        if ftte != 0:
            fttePas = fillfPassedfttePassage(w, boite, fttepass, T)
            endftteLineS = PboFillFTTeStocker(w, boite, ftt, fttePas, T)
            endftteLine = ftteFillIn(w, boitsF, boite, endftteLineS, 1)
            end = aroundTo(endftteLine, 12)
            p = fillPecPassage(w, x, 1, linestockstart, 0, T, p)
            p = fillPecPassage(w, x, Lin, ftteLine, Lin - 1, tubeRound(Lin), p)
            p = fillPecPassage(w, x, endftteLine, end + 1, endftteLine - 1, tubeRound(endftteLine), p)
        else:
            fillPecPassage(w, x, 1, linestockstart, 0, T, p)
            fillPecPassage(w, x, Lin, capacity + 1, Lin - 1, tubeRound(Lin), p)
    else:
        if ftte != 0:
            fttePas = librePassFTTEFill(w, boite, fttepass, T)
            endftteLineS = PboFillFTTeStocker(w, boite, ftt, fttePas, T)
            endftteLine = ftteFillIn(w, boites, boite, endftteLineS, 1)
            end = aroundTo(endftteLine, 12)
            p = libreFillIn(w, boite, 1, linestockstart, p)
            p = libreFillIn(w, boite, Lin, ftteLine, p)
            p = libreFillIn(w, boite, endftteLine, end + 1, p)
        else:
            p = libreFillIn(w, boite, 1, linestockstart, p)
            p = libreFillIn(w, boite, Lin, capacity + 1, p)

    extracablePECPBOFillIn(w, boites, boite, capacity + 1, p)


SROboite = getSroBoite()
print(SROboite)

# ############## start fill In the pds ##########################################################
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
    w = workbook.add_worksheet(str(boiteCode[b]))
    boite = boiteCode[b]
    func = boiteFunction[b]
    cable = boiteCable[b]
    capacity = getCapacity(cable)
    if func == 'PEC':
        boitePecFillIn(w, cable, boite, capacity, T)
    elif func == 'PEC-PBO' or func == 'BTI':
        boitePecPboFillIn(w, cable, boite, capacity, T)
    else:
        boitePboFillIn(w, cable, boite, capacity, T)
workbook.close()

# ################# some test for verification ##############################################
boite = 'BTI-21-011-075-1007'
cab = getCable(boite)
cap = getCapacity(cab)
index1 = boiteCode.index(boite)
# # index2 = boiteCode.index('PBO-21-011-076-3006')
# cable = getCable('PBO-21-011-076-2015')
# cap = getCapacity(cable)
print(nbf[index1], checkGlobalFtt(boite), getPassedFtte(boite, cap), getNumbrFu(getLastStartBoite(boite), 0),
      getNumbrFu(boite, 0), getStockStartLine(boite),
      getLastStartBoite(boite))

for r in listCasseteNotfound:
    print(r)
# ftte = getNumbrFu('PBO-21-011-076-2015', 0)
# index = boiteCode.index('PBO-21-011-076-3011')
# ref = boiteReference[index]
# fuUsed = checkGlobalFtt(getLastStartBoite('PBO-85-081-855-1033')) - checkGlobalFtt(
# getLastStartBoite('PBO-85-081-855-1032'))
# ftteLine = getFTTElineStart('PBO-85-081-855-1033') + checkGlobalFtt(getLastStartBoite('PBO-85-081-855-1033'))
# fuBoit = getNumbrFu('PBO-85-081-855-1033', 0) - checkGlobalFtt('PBO-85-081-855-1033')
# ftteBoite = getFTTEBoites('PBO-85-081-855-1033')
# print(ftteLine, ftteBoite, checkGlobalFtt('PBO-85-081-855-1033'), fuUsed)
