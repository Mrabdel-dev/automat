import operator

from dbfread import DBF
import xlsxwriter
import datetime

# date configuration
now = datetime.datetime.now()
date = now.strftime("%d/%m/%Y")
# ################## load the both file boite and cable in DBF format ###################################
cableTable = DBF('pdsInput/85_048_568_CABLE_OPTIQUE_A.dbf', load=True, encoding='iso-8859-1')
boiteTable = DBF('pdsInput/85_048_568_BOITE_OPTIQUE_A.dbf', load=True, encoding='iso-8859-1')
zaPboDbl = DBF('pdsInput/zapbodbl.dbf', load=True, encoding='iso-8859-1')
casseteTable = DBF('pdsInput/cassete file.dbf', load=True, encoding='iso-8859-1')
# ################### declare the excel pds file ###########################################################
workbook = xlsxwriter.Workbook('PDS/pds-85_048_568.xlsx')
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
        nbfu = getNumbrFu(b, 0)
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
        nbfu = getNumbrFu(b, 0)
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
            nbmrEpes = getNumbrFu(b, nbmrEpes)
        return nbmrEpes


# math function to major a number to a specific num
def aroundTo(x: int, num):
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
                    return aroundTo(fuFttE, 3)
                else:
                    fuFttE += n * 2
                    return aroundTo(fuFttE, 3)
            elif t == 'FTTH' and y == 'BATIMENT PUBLIC':
                fuFttE += n * 2
                return aroundTo(fuFttE, 3)
    else:
        return 0


# get the resulte ftte of all next coming boite
def checkGlobalFtt(bo):
    listBoit = getListComingBoite(bo)
    x = 0
    if len(listBoit) == 0:
        return checkFtt(bo)
    elif len(listBoit) == 1:
        x = checkFtt(bo)
        x += checkFtt(listBoit[0])
        return x
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


# function return where i should start write to write stocked state
def getStockStartLine(boite):
    fuUsed = getNumbrFu(getLastStartBoite(boite), 0)
    fuBoit = getNumbrFu(boite, 0)
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


# function return where i should start write ftte
def getFTTElineStart(boite):
    cable = getCable(boite)
    capacity = getCapacity(cable)
    line = capacity - aroundTo(checkGlobalFtt(boite), 12)
    return line + 1


# function to get boit that have ftte prise
def getFTTEBoites(boite):
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
        w.write(i + 1, 2, num, border)
        w.write(i + 1, 3, T, stringCassette(str(T)))
        if num % 12 == 0:
            if T == 24:
                T = 1
            else:
                T += 1
        w.write(i + 1, 4, num, stringCassette(str(num)))
        w.write(i + 1, 5, '', border)


# function to write next cable epesuree on the boite just for specific next boit
def fillInEpess(w: sheet, Lin, i, boite, T, N, k, size):
    cable = getCable(boite)
    capacity = getCapacity(cable)
    ftt = checkGlobalFtt(boite)
    funb = getNumbrFu(boite, 0)
    nbrEps = funb - ftt
    for j in range(0, nbrEps):
        n = 'CSE-' + str(N)
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
        if k > size:
            k = 1
            N += 1
        i += 1
        Lin += 1
    return Lin, k, N


# function to write  all next cable epesuree on the boite
def fillInAllCableEpess(w: sheet, nextBoite, boite, Lin):
    index = getcassteIndex(boite)
    size = tailleCassete[index]
    ftte = checkGlobalFtt(boite)
    N = aroundTo(ftte, size) / size + 1
    k = 1
    for b in nextBoite:
        x, k, N = fillInEpess(w, Lin, 0, b, 1, N, k, size)
        Lin = x

    return Lin


# function  to write the ftte nex cable
def ftteFillIn(w, Listboites, boite, startLin, T):
    index = getcassteIndex(boite)
    size = tailleCassete[index]
    k = 1
    N = 1
    for b in Listboites:
        i = 0
        ftteN = checkGlobalFtt(b)
        cable = getCable(b)
        capacity = getCapacity(cable)
        T = tubeRound(capacity - ftteN)
        for j in range(0, ftteN):
            w.write(startLin, 5, 'CSE-' + str(N).zfill(2), border)
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
            if k > size:
                k = 1
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


# function  to write the passage  next cable
def fillPecPassage(w, boite, startLine, endLine, i, T):
    cable = getCable(boite)
    cap = getCapacity(cable)
    for k in range(startLine, endLine):
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


# function to write stoker state
def PboFillStokker(w: sheet, boite, stokker, Lin, T=1):
    i = Lin - 1
    index = getcassteIndex(boite)
    N = nbrCassete[index]
    size = tailleCassete[index]
    k = 1
    for s in range(0, stokker):
        n = 'CSE-' + str(N)
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
        if k > size:
            N = N - 1
            k = 1
        Lin += 1
        i += 1
    return Lin


# function to write epssure state for PEC-PBO
def PboFillEpes(w: sheet, boites, boite, Lin, i, T=1):
    index = getcassteIndex(boite)
    size = tailleCassete[index]
    nbrCas = nbrCassete[index]
    ftte = checkGlobalFtt(boite)
    N = aroundTo(ftte, size) / size + 1
    k = 1
    for s in boites:
        x, k, N = fillInEpess(w, Lin, i, s, 1, N, k, size)
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


# function to write the libre state for next cable
def libreFillIn(w: sheet, boit, startLine, endLine, T=1):
    i = 1
    for k in range(startLine, endLine):
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


def extracableFillIn(w: sheet, cable, cap, extarline, startLine, funm):
    i = funm
    T = tubeRound(funm)
    for e in range(0, extarline):
        w.write(startLine, 0, '', border)
        w.write(startLine, 1, '', border)
        w.write(startLine, 2, '', border)
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
    return startLine


# function to write all extract libre cable need for sorted cable
def extracablePECPBOFillIn(w: sheet, boites, boite, startLine):
    y = getBoitePassage(boite)
    index1 = boiteCode.index(boite)
    func = boiteFunction[index1]
    k = getFTTEBoites(boite)
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
                fuN = getNumbrFu(b, 0)
                cable = getCable(b)
                cap = getCapacity(cable)
                extraN = cap - fuN
                startLine = extracableFillIn(w, cable, cap, extraN, startLine, fuN)

            else:
                cable = getCable(b)
                cap = getCapacity(cable)
                fu = getNumbrFu(b, 0)
                total = fuNumbr1 - fu + 1
                if func == 'PEC':
                    start = getLastStartBoite(boite)
                    if start == boite:
                        Lin = 1
                    else:
                        Lin = getNumbrFu(getLastStartBoite(start), 0)
                    startLine = extracableFillIn(w, cable, cap, fuNumbr1, startLine, Lin)
                else:
                    Lin = getStockStartLine(boite)
                    startLine = extracableFillIn(w, cable, cap, total, startLine, Lin)


        else:
            if y != b:
                cable = getCable(b)
                cap = getCapacity(cable)
                ftte = checkGlobalFtt(b)
                nbfu = getNumbrFu(b, 0) - ftte
                extraN2 = cap - aroundTo(ftte, 12) - getNumbrFu(b, 0) + ftte
                startLine = extracableFillIn(w, cable, cap, extraN2, startLine, nbfu)
                extraN = aroundTo(ftte, 12) - ftte
                tt = cap - aroundTo(ftte, 12) + ftte
                startLine = extracableFillIn(w, cable, cap, extraN, startLine, tt)

            else:
                cable = getCable(b)
                cap = getCapacity(cable)
                ftte = checkGlobalFtt(b)
                if func == 'PEC':
                    start = getLastStartBoite(boite)
                    if start == boite:
                        Lin = 1
                    else:
                        Lin = getNumbrFu(getLastStartBoite(start), 0)
                    startLine = extracableFillIn(w, cable, cap, fuNumbr1, startLine, Lin)
                else:
                    Lin = getStockStartLine(boite)
                    fu = getNumbrFu(b, 0)
                    total = fuNumbr1 - fu + 1
                    startLine = extracableFillIn(w, cable, cap, total, startLine, Lin)
                    # lin = Lin + getNumbrFu(getLastStartBoite(b), 0)-ftte
                    # startLine = extracableFillIn(w, cable, cap, fuNumbr1-ftte, startLine, Lin)


def getcassteIndex(boite):
    index = boiteCode.index(boite)
    ref = boiteReference[index]
    try:
        indexCass = reference.index(ref)
        return indexCass
    except ValueError:
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
    indexCass = getcassteIndex(boite)
    fuNumber = getNumbrFu(boite, 0)
    ftte = checkGlobalFtt(boite)
    ftteLine = getFTTElineStart(boite)
    Test = fuNumber + aroundTo(ftte, 12)
    if Test > capacity:
        print("Erouuuuuuuuur cable capacity not enough")
    cableBaseInfo(w, cable, capacity, T)
    nextBoits = getListComingBoitePEC(boite)
    boites = getListComingBoite(boite)
    start = getLastStartBoite(boite)
    print(start)
    if start == boite:
        Lin = 1
    else:
        Lin = getNumbrFu(getLastStartBoite(start), 0)
        libreFillIn(w, boite, 1, Lin, T)
    if len(boites) < 2:
        Lin = fillInAllCableEpess(w, boites, boite, Lin)
        endFTTLine = ftteFillIn(w, boites, ftteLine, T)
    else:
        Lin = fillInAllCableEpess(w, nextBoits, boite, Lin)
        endFTTLine = ftteFillIn(w, nextBoits, boite, ftteLine, T)

    x = getBoitePassage(boite)
    if x is not None:
        fillPecPassage(w, x, Lin, ftteLine, Lin - 1, tubeRound(Lin))
        fillPecPassage(w, x, endFTTLine, capacity + 1, endFTTLine - 1, tubeRound(endFTTLine))
    else:
        libreFillIn(w, boite, Lin, ftteLine, T)
        libreFillIn(w, boite, endFTTLine, capacity + 1, T)
    extracablePECPBOFillIn(w, boites, boite, capacity + 1)


# <--################### PEC-PBO function ##########################-->
def boitePecPboFillIn(w: sheet, cable, boite, capacity, T):
    endftteLine = 0
    ftteLine = 0
    baseSheet(w, boite)
    indexCass = getcassteIndex(boite)
    cableBaseInfo(w, cable, capacity, T)
    linestockstart = getStockStartLine(boite) + 1
    index = boiteCode.index(boite)
    ftte = checkGlobalFtt(boite)
    stoker = nbf[index] - checkFtt(boite)
    Lin = PboFillStokker(w, boite, stoker, linestockstart, 1)
    boites = getListComingBoitePEC(boite)
    Lin = PboFillEpes(w, boites, boite, Lin, 0, 1)
    ftteBoite = getFTTEBoites(boite)
    k = len(ftteBoite)
    if k > 0:
        ftteLine = getFTTElineStart(boite)
        boites = getListComingBoite(boite)
        endftteLine = ftteFillIn(w, boites, boite, ftteLine, 1)
    else:
        ftteLine = getFTTElineStart(boite)
        endftteLine = PboFillStokker(w, boite, ftte, ftteLine, 1)
    x = getBoitePassage(boite)
    if x is not None:
        fillPecPassage(w, x, 1, linestockstart, 0, T)
        fillPecPassage(w, x, Lin, ftteLine, Lin - 1, tubeRound(Lin))
        fillPecPassage(w, x, endftteLine, capacity + 1, endftteLine - 1, tubeRound(endftteLine))
    else:
        libreFillIn(w, boite, 1, linestockstart, 1)
        libreFillIn(w, boite, Lin, ftteLine, T)
        libreFillIn(w, boite, endftteLine, capacity + 1, T)
    boites = getListComingBoite(boite)
    extracablePECPBOFillIn(w, boites, boite, capacity + 1)


# <--################### PBO function #############################-->
def boitePboFillIn(w: sheet, cable, boite, capacity, T):
    baseSheet(w, boite)
    indexCass = getcassteIndex(boite)
    cableBaseInfo(w, cable, capacity, T)
    linestockstart = getStockStartLine(boite) + 1
    index = boiteCode.index(boite)
    ftte = checkGlobalFtt(boite)
    stoker = nbf[index] - checkFtt(boite)
    Lin = PboFillStokker(w, boite, stoker, linestockstart, 1)
    ftteLine = getFTTElineStart(boite)
    boites = getFTTEBoites(boite)
    if len(boites) < 1:
        endftteLine = PboFillStokker(w, boite, ftte, ftteLine, 1)
    else:
        endftteLine = ftteLine
    x = getBoitePassage(boite)
    if x is not None:
        fillPecPassage(w, x, 1, linestockstart, 0, T)
        fillPecPassage(w, x, Lin, ftteLine, Lin - 1, tubeRound(Lin))
        fillPecPassage(w, x, endftteLine, capacity + 1, endftteLine - 1, tubeRound(endftteLine))
    else:
        libreFillIn(w, boite, 1, linestockstart, 1)
        libreFillIn(w, boite, Lin, ftteLine, T)
        libreFillIn(w, boite, endftteLine, capacity + 1, T)
    boites = getListComingBoite(boite)
    extracablePECPBOFillIn(w, boites, boite, capacity + 1)


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
    elif func == 'PEC-PBO':
        boitePecPboFillIn(w, cable, boite, capacity, T)
    else:
        boitePboFillIn(w, cable, boite, capacity, T)
workbook.close()

# ################# some test for verification ##############################################
# index1 = boiteCode.index('PBO-21-011-076-2015')
# # index2 = boiteCode.index('PBO-21-011-076-3006')
# cable = getCable('PBO-21-011-076-2015')
# cap = getCapacity(cable)
# print(nbf[index1])
# # print(getFTTEBoites('PBO-21-011-076-3035'))
# x = checkGlobalFtt('PBO-21-011-076-2015')
# print(x)
# ftte = getNumbrFu('PBO-21-011-076-2015', 0)
# index = boiteCode.index('PBO-21-011-076-3011')
# ref = boiteReference[index]
