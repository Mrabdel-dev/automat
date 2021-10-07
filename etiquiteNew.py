from dbfread import DBF
import xlsxwriter
import datetime

# date configuration
now = datetime.datetime.now()
date = now.strftime("%m/%Y")
# ################## load the both file boite and cable in DBF format ###################################
cableTable = DBF('etiqueteInputs/21_011_078_CABLE_OPTIQUE_B_AI.dbf', load=True, encoding='iso-8859-1')
boiteTable = DBF('etiqueteInputs/21_011_078_BOITE_OPTIQUE_B.dbf', load=True, encoding='iso-8859-1')
pointTechTable = DBF('etiqueteInputs/21_011_078_POINT_TECHNIQUE_B.dbf', load=True, encoding='iso-8859-1')
joinTable = DBF('etiqueteInputs/joinCablePT-078.dbf', load=True, encoding='iso-8859-1')
fciTable = DBF('etiqueteInputs/fCI-078.dbf', load=True, encoding='iso-8859-1')
Fibre = "ALTITUDE FIBRE 21"
propFibre = "ALTI"
sro = 'SRO-21-011-078'
# ################### declare the excel pds file ###########################################################
workbook = xlsxwriter.Workbook(f'Etiquette/{sro}-DETAIL-ETIQUETTE.xlsx')
workbook1 = xlsxwriter.Workbook(f'Etiquette/{sro}-ETIQUETTE.xlsx')
totaleSheet = workbook1.add_worksheet("EtiquettePrintedFile")

# ############### define the character and style of cell inside excel ################"0
border = workbook.add_format({"border": 1})
bold = workbook.add_format({'bold': True, "border": 1})
border1 = workbook1.add_format({"border": 1})
bold1 = workbook1.add_format({'bold': True, "border": 1})
# charge the name of all filed in tables
filedCableNam = cableTable.field_names
filedBoiteNam = boiteTable.field_names
filedFciNam = fciTable.field_names
boiteLen = len(boiteTable)
cableLen = len(cableTable)
pointlen = len(pointTechTable)
fcilen = len(fciTable)
joinlen = len(joinTable)
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
    try:
        fciNom.append(fciTable.records[f]['POTEAU_CHA'])
    except KeyError:
        fciNom.append(fciTable.records[f]['N__'])

    fciCode.append(fciTable.records[f]['FCI'])

allcable = []
cablePTCode = []
typeFonc = []
typeStruc = []
cablePTProp = []
for k in range(0, joinlen):
    allcable.append(joinTable.records[k]['NOM'])
    cablePTCode.append(joinTable.records[k]['NOM_2'])
    typeFonc.append(joinTable.records[k]['TYPE_FONC_'])
    typeStruc.append(joinTable.records[k]['TYPE_STR_1'])
    cablePTProp.append(joinTable.records[k]['GESTIONN_1'])
# ###################### define  the base header ##############################
sheet = xlsxwriter.worksheet.Worksheet

listErour = []


# ################ vender ###########
def venderBaseHeader():
    pass


# ################# normale ############
def baseHeader(w: sheet):
    w.write('A1', "CODE POINT TECHNIQUE", bold)
    w.write('B1', 'NB ETIQUETTE', bold)
    w.write('C1', 'COULEUR_ETIQUETTE', bold)
    w.write('D1', 'LIGNE 1', bold)
    w.write('E1', 'LIGNE 2', bold)
    w.write('F1', 'LIGNE 3', bold)
    w.write('G1', 'LIGNE 4', bold)


def baseHeader1(w: sheet):
    w.write('A1', "CODE POINT TECHNIQUE", bold1)
    w.write('B1', 'NB ETIQUETTE', bold1)
    w.write('C1', 'COULEUR_ETIQUETTE', bold1)
    w.write('D1', 'LIGNE 1', bold1)
    w.write('E1', 'LIGNE 2', bold1)
    w.write('F1', 'LIGNE 3', bold1)
    w.write('G1', 'LIGNE 4', bold1)


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


def getpointIndex(boite):
    index = boiteCode.index(boite)
    idPrent = boiteIdParent[index]
    indexPoint = pointCode.index(idPrent)
    return indexPoint


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


def fillInAllTable(cable, pointCode, index):
    allcable.append(cable)
    cablePTCode.append(pointCode)
    typeFonc.append(pointFonc[index])
    typeStruc.append(pointStruc[index])
    cablePTProp.append(pointPrp[index])


# ######### fill in ############
def boiteEtiqueteFill(boites, k, totale: sheet):
    w = workbook.add_worksheet("Etiquette Boite")
    baseHeader(w)

    lin = 2
    for b in boites:
        if not str(b).startswith('SRO'):
            pointTech = getPointTech(b)
            index = getpointIndex(b)
            prop = str(pointPrp[index])
            fcicode = getFci(pointTech)
            if fcicode is not None:
                w.write('A' + str(lin), getPointCode(b), border)
                w.write('B' + str(lin), '1', border)
                w.write('C' + str(lin), 'BLANC', border)
                w.write('D' + str(lin), Fibre, border)
                w.write('E' + str(lin), b, border)
                w.write('F' + str(lin), str(fcicode) + "-" + str(date), border)
                w.write('G' + str(lin), '', border)
                # ############################
                totale.write('A' + str(k), pointTech, border1)
                totale.write('B' + str(k), '1', border1)
                totale.write('C' + str(k), 'BLANC', border1)
                totale.write('D' + str(k), Fibre, border1)
                totale.write('E' + str(k), b, border1)
                totale.write('F' + str(k), str(fcicode) + "-" + str(date), border1)
                totale.write('G' + str(k), '', border1)
                lin += 1
                k += 1
            else:
                if prop.startswith(propFibre) or prop.startswith("ENE"):
                    fcicode = sro
                else:
                    fcicode = "  "
                w.write('A' + str(lin), getPointCode(b), border)
                w.write('B' + str(lin), '1', border)
                w.write('C' + str(lin), 'BLANC', border)
                w.write('D' + str(lin), Fibre, border)
                w.write('E' + str(lin), b, border)
                w.write('F' + str(lin), fcicode + "-" + str(date), border)
                w.write('G' + str(lin), '', border)
                # ############################
                totale.write('A' + str(k), pointTech, border1)
                totale.write('B' + str(k), '1', border1)
                totale.write('C' + str(k), 'BLANC', border1)
                totale.write('D' + str(k), Fibre, border1)
                totale.write('E' + str(k), b, border1)
                totale.write('F' + str(k), fcicode + "-" + str(date), border1)
                totale.write('G' + str(k), '', border1)
                lin += 1
                k += 1


def pointEtiqueteFill(points, k, totale: sheet):
    po = workbook.add_worksheet("Etiquette PT FIBRE 21")
    baseHeader(po)
    lin = 2
    for p in points:
        prop = getProp(p)
        if prop.startswith(propFibre):
            po.write('A' + str(lin), p, border)
            po.write('B' + str(lin), '1', border)
            po.write('C' + str(lin), 'BLANC', border)
            po.write('D' + str(lin), prop, border)
            po.write('E' + str(lin), p, border)
            po.write('F' + str(lin), sro + "-" + str(date), border)
            po.write('G' + str(lin), '', border)
            # ############################
            totale.write('A' + str(k), p, border1)
            totale.write('B' + str(k), '1', border1)
            totale.write('C' + str(k), 'BLANC', border1)
            totale.write('D' + str(k), prop, border1)
            totale.write('E' + str(k), p, border1)
            totale.write('F' + str(k), sro + "-" + str(date), border1)
            totale.write('G' + str(k), '', border1)
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
        cap = getCapacity(cable)
        point = cablePTCode[i]
        nm = getNomPT(point)
        fci = getFci(nm)
        typef = typeFonc[i]
        typeStr = typeStruc[i]
        if typeStr == 'CHAMBRE' or point.startswith('SRO'):
            N = 2
        cbPro = cablePTProp[i]
        if fci is None and cbPro != 'ORANGE':
            fci = sro
        if typef == 'TIRAGE' and (typeStr == 'APPUI' or typeStr == 'ANCRAGE FACADE') and (
                cbPro == 'ORANGE' or cbPro == 'ENEDIS' or cbPro == 'PROPRIETAIRE PRIVE'):
            continue
        else:

            po.write('A' + str(lin), point, border)
            po.write('B' + str(lin), N, border)
            po.write('C' + str(lin), 'BLANC', border)
            po.write('D' + str(lin), Fibre, border)
            po.write('E' + str(lin), str(cable) + "-" + str(cap) + " FO", border)
            po.write('F' + str(lin), str(fci) + "-" + str(date), border)
            po.write('G' + str(lin), '', border)
            # ############################
            totale.write('A' + str(k), point, border1)
            totale.write('B' + str(k), N, border1)
            totale.write('C' + str(k), 'BLANC', border1)
            totale.write('D' + str(k), Fibre, border1)
            totale.write('E' + str(k), str(cable) + "-" + str(cap) + " FO", border1)
            totale.write('F' + str(k), str(fci) + "-" + str(date), border1)
            totale.write('G' + str(k), '', border1)
            lin += 1
            k += 1
    return k


def etiquettePtOrangeFill(cables, totale: sheet):
    co = workbook.add_worksheet("Etiquette Poteau ORANGE")
    baseHeader(co)
    baseHeader(totale)
    x = len(cables)
    lin = 2
    for i in range(0, x):
        N = 1
        cable = allcable[i]
        cap = getCapacity(cable)
        point = cablePTCode[i]
        nm = getNomPT(point)
        fci = getFci(nm)
        if fci is None:
            fci = ''
        typeStr = typeStruc[i]
        if typeStr == 'CHAMBRE':
            N = 2
        cbPro = cablePTProp[i]
        if typeStr == 'APPUI' and cbPro == 'ORANGE':
            co.write('A' + str(lin), point, border)
            co.write('B' + str(lin), N, border)
            co.write('C' + str(lin), 'BLANC', border)
            co.write('D' + str(lin), Fibre, border)
            co.write('E' + str(lin), str(cable) + "-" + str(cap) + " FO", border)
            co.write('F' + str(lin), str(fci) + "-" + str(date), border)
            co.write('G' + str(lin), '', border)
            # ############################
            totale.write('A' + str(lin), point, border1)
            totale.write('B' + str(lin), N, border1)
            totale.write('C' + str(lin), 'BLANC', border1)
            totale.write('D' + str(lin), Fibre, border1)
            totale.write('E' + str(lin), str(cable) + "-" + str(cap) + " FO", border1)
            totale.write('F' + str(lin), str(fci) + "-" + str(date), border1)
            totale.write('G' + str(lin), '', border1)
            lin += 1
    k = lin
    return k


# #######################################################################
k = etiquettePtOrangeFill(allcable, totaleSheet)
k = pointEtiqueteFill(pointCode, k, totaleSheet)
k = cableEtiqueteFill(allcable, k, totaleSheet)
boiteEtiqueteFill(boiteCode, k, totaleSheet)
print('#' * 45)
for i in range(0, len(allcable)):
    print(allcable[i], cablePTCode[i], typeFonc[i], typeStruc[i], cablePTProp[i])
print('#' * 45)
workbook.close()
workbook1.close()
for c in listErour:
    print(c)
