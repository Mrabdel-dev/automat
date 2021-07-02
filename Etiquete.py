from dbfread import DBF
import xlsxwriter
import datetime

# date configuration
now = datetime.datetime.now()
date = now.strftime("%m/%Y")
# ################## load the both file boite and cable in DBF format ###################################
cableTable = DBF('etiqueteInputs/21_017_103_CABLE_OPTIQUE_A.dbf', load=True, encoding='iso-8859-1')
boiteTable = DBF('etiqueteInputs/21_017_103_BOITE_OPTIQUE_A.dbf', load=True, encoding='iso-8859-1')
pointTechTable = DBF('etiqueteInputs/21_017_103_POINT_TECHNIQUE_A.dbf', load=True, encoding='iso-8859-1')
supportTable = DBF('etiqueteInputs/21_017_103_SUPPORT_A.dbf', load=True, encoding='iso-8859-1')
fciTable = DBF('etiqueteInputs/FCI_103.dbf', load=True, encoding='iso-8859-1')

# ################### declare the excel pds file ###########################################################
workbook = xlsxwriter.Workbook('Etiquette/etiquetteDetail103.xlsx')
totaleSheet = workbook.add_worksheet("EtiquettePrintedFile")
# ############### define the character and style of cell inside excel ################"
border = workbook.add_format({"border": 1})
header = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#C4E5F7'})
# charge the name of all filed in tables
filedCableNam = cableTable.field_names
filedBoiteNam = boiteTable.field_names
filedFciNam = fciTable.field_names
boiteLen = len(boiteTable)
cableLen = len(cableTable)
pointlen = len(pointTechTable)
fcilen = len(fciTable)
supplen = len(supportTable)
# #######################declare the table that i need te full#############################################
# FROM THE BOITE OPTIQUE
boiteCode = []  # name of the boite
boiteIdParent = []  # AMOUNT CABLE
for j in range(0, boiteLen):
    boiteCode.append(boiteTable.records[j]['CODE'])
    boiteIdParent.append(boiteTable.records[j]['ID_PARENT'])
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
for s in range(0, supplen):
    suppAmount.append(supportTable.records[s]['AMONT'])
    suppAval.append(supportTable.records[s]['AVAL'])
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
# FROM THE FCI 
fciNom = []
fciCode = []
for f in range(0, fcilen):
    fciNom.append(fciTable.records[f]['POTEAU_CHA'])
    fciCode.append(fciTable.records[f]['FCI'])

# ###################### define  the base header ##############################
sheet = xlsxwriter.worksheet.Worksheet


# ################ vender ###########
def venderBaseHeader():
    pass


# ################# normale ############
def baseHeader(w: sheet):
    w.write('A1', "CODE POINT TECHNIQUE", header)
    w.write('B1', 'NB ETIQUETTE', header)
    w.write('C1', 'COULEUR_ETIQUETTE', header)
    w.write('D1', 'LIGNE 1', header)
    w.write('E1', 'LIGNE 2', header)
    w.write('F1', 'LIGNE 3', header)
    w.write('G1', 'LIGNE 4', header)


# #################################### function  ########################
def getPointCode(boite):
    index = boiteCode.index(boite)
    idPrent = boiteIdParent[index]
    return idPrent


def getCapacity(cable):
    index = cableName.index(cable)
    cap = cableCapacity[index]
    return cap


def getNomPT(code):
    try:
        index = pointCode.index(code)
        fcNom = pointNom[index]
        fcNom = fcNom[0:6] + fcNom[6:].lstrip("0")
        return fcNom
    except ValueError:
        return None


def getPointTech(boite):
    index = boiteCode.index(boite)
    idPrent = boiteIdParent[index]
    indexPoint = pointCode.index(idPrent)
    pointTech = pointNom[indexPoint]
    pointTech = pointTech[0:6] + pointTech[6:].lstrip("0")
    return pointTech


def getFci(pointTech):
    try:
        index = fciNom.index(pointTech)
        fcicode = fciCode[index]
        return fcicode
    except ValueError:
        return None


def getProp(point):
    index = pointCode.index(point)
    prop = str(pointPrp[index])
    return prop


def getCablePointTechStart(cable):
    index = cableName.index(cable)
    boite = cableOrigin[index]
    pointTechCode = getPointCode(boite)
    return pointTechCode


def getCablePointTechEnd(cable):
    index = cableName.index(cable)
    boite = cableExtremity[index]
    pointTechCode = getPointCode(boite)
    return pointTechCode


def duplicates(lst, item):
    return [i for i, x in enumerate(lst) if x == item]


allcable = []
cablePTCode = []
typeFonc = []
typeStruc = []
cablePTProp = []


def fillInAllTable(cable, pointCode, index):
    allcable.append(cable)
    cablePTCode.append(pointCode)
    typeFonc.append(pointFonc[index])
    typeStruc.append(pointStruc[index])
    cablePTProp.append(pointPrp[index])


def getAval(point):
    index = suppAmount.index(point)
    return suppAval[index]


def getAmount(point):
    index = suppAval.index(point)
    return suppAmount[index]


def checkAvalway( cablepointStart, cablepointEnd, test):
    start = cablepointEnd
    k = test
    try:
        while test and k:
            if cablepointStart == start:
                test = False
                break
            else:
                start = getAmount(start)
                listDb = duplicates(suppAmount, start)
                if len(listDb) > 1:
                    if cablepointStart == start:
                        test = False
                    k = False
                    break

        return test
    except ValueError:
        print(start)
        return True


def fillAvalWay(c, cablepointStart, cablepointEnd, test):
    start = cablepointEnd
    k = test
    try:
        while test and k:
            if cablepointStart == start:
                test = False
                break
            else:
                start = getAmount(start)
                listDb = duplicates(suppAmount, start)
                if len(listDb) > 1:
                    if cablepointStart == start:
                        index = pointCode.index(cablepointStart)
                        fillInAllTable(c, cablepointStart, index)
                        test = False
                    k = False
                    break
                index = pointCode.index(start)
                fillInAllTable(c, start, index)
        return test
    except ValueError:
        print(start)
        return True


def fillInTable(c):
    test = True
    cablePointStart = getCablePointTechStart(c)
    cablePointEnd = getCablePointTechEnd(c)
    index = pointCode.index(cablePointEnd)
    fillInAllTable(c, cablePointEnd, index)
    test = checkAvalway(cablePointStart, cablePointEnd, test)
    if test:
        index = pointCode.index(cablePointStart)
        fillInAllTable(c, cablePointStart, index)
    else :
        test = True
        test = fillAvalWay(c,cablePointStart, cablePointEnd, test)
    print(test)
    print(cablePointStart)
    try:
        nextstart = getAval(cablePointStart)
        print('#', nextstart)
        duplistart = duplicates(suppAval, nextstart)
    except ValueError:

        nextstart = getAmount(cablePointStart)
        duplistart = duplicates(suppAmount, nextstart)
        print(c, cablePointStart, nextstart, duplicates(suppAmount, nextstart))
    try:
        nextend = getAmount(cablePointEnd)
        dupliend = duplicates(suppAmount, nextend)
        print(c, nextend, len(dupliend), len(duplistart))
    except ValueError:
        nextend = getAval(cablePointEnd)
        print(c, cablePointEnd, nextend)
        dupliend = duplicates(suppAval, nextend)
    k = 0

    while test:

        try:
            if len(dupliend) <= 1 and len(duplistart) <= 1:
                if nextend == nextstart:
                    index = pointCode.index(nextend)
                    fillInAllTable(c, nextend, index)
                    test = False

                elif nextstart == cablePointEnd or nextend == cablePointStart:
                    test = False

                else:
                    index = pointCode.index(nextend)
                    fillInAllTable(c, nextend, index)
                    index = pointCode.index(nextstart)
                    fillInAllTable(c, nextstart, index)
                    nextstart = getAval(nextstart)
                    nextend = getAmount(nextend)
            elif len(dupliend) < 2 and len(duplistart) > 1:
                if nextend == nextstart:
                    index = pointCode.index(nextend)
                    fillInAllTable(c, nextend, index)
                    test = False
                    break
                elif nextstart == cablePointEnd or nextend == cablePointStart:
                    test = False
                    break
                else:
                    index = pointCode.index(nextend)
                    fillInAllTable(c, nextend, index)
                    nextend = getAmount(nextend)

            elif len(duplistart) < 2 and len(dupliend) > 1:
                if nextend == nextstart:
                    index = pointCode.index(nextend)
                    fillInAllTable(c, nextend, index)
                    test = False
                    break
                elif nextstart == cablePointEnd or nextend == cablePointStart:
                    test = False
                    break
                else:
                    index = pointCode.index(nextstart)
                    fillInAllTable(c, nextstart, index)
                    nextstart = getAval(nextstart)



            else:
                if nextend == nextstart:
                    index = pointCode.index(nextend)
                    fillInAllTable(c, nextend, index)

                test = False
            dupliend = duplicates(suppAmount, nextend)
            duplistart = duplicates(suppAval, nextstart)
        except ValueError:
            print(c)
            print(' start', cablePointStart)
            print('end', cablePointEnd)
            print('next start', nextstart)
            print('next end', nextend)
            test = False


def createTablesBase(cables):
    for c in cables:
        fillInTable(c)



createTablesBase(cableName)


# ############################# fill in function #################
def boiteEtiqueteFill(boites, totale: sheet):
    w = workbook.add_worksheet("EtiquetteBoite")
    baseHeader(w)
    baseHeader(totale)
    lin = 2
    for b in boites:
        pointTech = getPointTech(b)
        fcicode = getFci(pointTech)
        if fcicode is not None:
            w.write('A' + str(lin), getPointCode(b), border)
            w.write('B' + str(lin), '1', border)
            w.write('C' + str(lin), 'BLANC', border)
            w.write('D' + str(lin), 'ALTITUDE FIBRE 21', border)
            w.write('E' + str(lin), b, border)
            w.write('F' + str(lin), str(fcicode) + str(date), border)
            w.write('G' + str(lin), '', border)
            # ############################
            totale.write('A' + str(lin), pointTech, border)
            totale.write('B' + str(lin), '1', border)
            totale.write('C' + str(lin), 'BLANC', border)
            totale.write('D' + str(lin), 'ALTITUDE FIBRE 21', border)
            totale.write('E' + str(lin), b, border)
            totale.write('F' + str(lin), str(fcicode) + " " + str(date), border)
            totale.write('G' + str(lin), '', border)
            lin += 1
            k = lin
    return k


def pointEtiqueteFill(points, k, totale: sheet):
    po = workbook.add_worksheet("Etiquette PT")
    baseHeader(po)
    lin = 2
    for p in points:
        prop = getProp(p)
        if prop.startswith("ALTITUDE"):
            po.write('A' + str(lin), p, border)
            po.write('B' + str(lin), '1', border)
            po.write('C' + str(lin), 'BLANC', border)
            po.write('D' + str(lin), prop, border)
            po.write('E' + str(lin), p, border)
            po.write('F' + str(lin), "                  " + str(date), border)
            po.write('G' + str(lin), '', border)
            # ############################
            totale.write('A' + str(k), p, border)
            totale.write('B' + str(k), '1', border)
            totale.write('C' + str(k), 'BLANC', border)
            totale.write('D' + str(k), prop, border)
            totale.write('E' + str(k), p, border)
            totale.write('F' + str(k), "                  " + str(date), border)
            totale.write('G' + str(k), '', border)
            lin += 1
            k += 1
    return k


def cableEtiqueteFill(cables, k, totale: sheet):
    po = workbook.add_worksheet("Etiquette Cable")
    baseHeader(po)

    x = len(cables)
    lin = 2
    for i in range(0, x):
        N = 1
        cable = allcable[i]
        y = str(cable)[3:14]
        cap = getCapacity(cable)
        point = cablePTCode[i]
        nm = getNomPT(point)
        fci = getFci(nm)
        if fci is None:
            fci = 'SRO' + y
        typef = typeFonc[i]
        typeStr = typeStruc[i]
        if typeStr == 'CHAMBRE':
            N = 2
        cbPro = cablePTProp[i]
        if typef == 'TIRAGE' and (typeStr == 'APPUI' or typeStr == 'ANCRAGE FACADE') and (
                cbPro == 'ORANGE' or cbPro == 'ENEDIS' or cbPro == 'PROPRIETAIRE PRIVE'):
            continue
        else:

            po.write('A' + str(lin), point, border)
            po.write('B' + str(lin), N, border)
            po.write('C' + str(lin), 'BLANC', border)
            po.write('D' + str(lin), "ALTITUDE FIBRE 21", border)
            po.write('E' + str(lin), str(cable) + "-" + str(cap) + " FO", border)
            po.write('F' + str(lin), str(fci) + "-" + str(date), border)
            po.write('G' + str(lin), '', border)
            # ############################
            totale.write('A' + str(k), point, border)
            totale.write('B' + str(k), N, border)
            totale.write('C' + str(k), 'BLANC', border)
            totale.write('D' + str(k), "ALTITUDE FIBRE 21", border)
            totale.write('E' + str(k), str(cable) + "-" + str(cap) + " FO", border)
            totale.write('F' + str(k), str(fci) + "-" + str(date), border)
            totale.write('G' + str(k), '', border)
            lin += 1
            k += 1
    return k


def etiquettePtOrangeFill(cables, k, totale: sheet):
    co = workbook.add_worksheet("Etiquette PT Orange")
    baseHeader(co)
    x = len(cables)
    lin = 2
    for i in range(0, x):
        N = 1
        cable = allcable[i]
        cap = getCapacity(cable)
        point = cablePTCode[i]
        nm = getNomPT(point)
        fci = getFci(nm)
        typeStr = typeStruc[i]
        if typeStr == 'CHAMBRE':
            N = 2
        cbPro = cablePTProp[i]
        if typeStr == 'APPUI' and cbPro == 'ORANGE':
            co.write('A' + str(lin), point, border)
            co.write('B' + str(lin), N, border)
            co.write('C' + str(lin), 'BLANC', border)
            co.write('D' + str(lin), "ALTITUDE FIBRE 21", border)
            co.write('E' + str(lin), str(cable) + "-" + str(cap) + " FO", border)
            co.write('F' + str(lin), str(fci) + "-" + str(date), border)
            co.write('G' + str(lin), '', border)
            # ############################
            totale.write('A' + str(k), point, border)
            totale.write('B' + str(k), N, border)
            totale.write('C' + str(k), 'BLANC', border)
            totale.write('D' + str(k), "ALTITUDE FIBRE 21", border)
            totale.write('E' + str(k), str(cable) + "-" + str(cap) + " FO", border)
            totale.write('F' + str(k), str(fci) + "-" + str(date), border)
            totale.write('G' + str(k), '', border)
            lin += 1
            k += 1


# #######################################################################
k = boiteEtiqueteFill(boiteCode, totaleSheet)
k = pointEtiqueteFill(pointCode, k, totaleSheet)
k = cableEtiqueteFill(allcable, k, totaleSheet)
etiquettePtOrangeFill(allcable, k, totaleSheet)
print('#' * 45)

for i in range(590, 610):
    print(allcable[i], "    ", cablePTCode[i], "  ", typeFonc[i], " ", typeStruc[i], " ", cablePTProp[i])
workbook.close()


