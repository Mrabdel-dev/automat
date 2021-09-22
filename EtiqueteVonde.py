from dbfread import DBF
import xlsxwriter
import datetime

# date configuration
now = datetime.datetime.now()
date = now.strftime("%m/%Y")
# ################## load the both file boite and cable in DBF format ###################################
boiteTable = DBF('vonder/85_048_570_BOITE_OPTIQUE_A2.dbf', load=True, encoding='iso-8859-1')
pointTechTable = DBF('vonder/85_048_570_POINT_TECHNIQUE_A2.dbf', load=True, encoding='iso-8859-1')
joinTable = DBF('vonder/joinCablePT-570.dbf', load=True, encoding='iso-8859-1')
fciTable = DBF('vonder/fCI-570.dbf', load=True, encoding='iso-8859-1')
codeTable = DBF('vonder/code_affaire.dbf', load=True, encoding='iso-8859-1')
sro = 'SRO-85-048-570'
# ################### declare the excel pds file ###########################################################
workbook = xlsxwriter.Workbook(f'Etiquette/{sro}-ETIQUETTE.xlsx')
totaleSheet = workbook.add_worksheet("ETIQUETTE")

# ############### define the character and style of cell inside excel ################"0
border = workbook.add_format({"border": 1})
header = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#DADADA'})
yellow = workbook.add_format({'bg_color': '#FFFF00'})
green = workbook.add_format({'bg_color': '#7CFC00'})
bleu = workbook.add_format({'bg_color': '#0F056B'})
white = workbook.add_format({'bg_color': '#FDF6F6'})

# charge the name of all filed in tables
filedCableNam = codeTable.field_names
filedBoiteNam = boiteTable.field_names
filedFciNam = fciTable.field_names
boiteLen = len(boiteTable)
cableLen = len(codeTable)
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
namEned = []  # NAME OF the poto endis
codeEned = []  # code d'affire

for i in range(0, cableLen):
    namEned.append(codeTable.records[i]['NOM'])
    codeEned.append(codeTable.records[i]['N_'])

# FROM THE SUPPORT

# FROM THE TECHNIC POINT
pointNom = []
pointCode = []
pointFonc = []
pointStruc = []
rempli = []
pointPrp = []
for p in range(0, pointlen):
    pointNom.append(pointTechTable.records[p]['NOM'])
    pointCode.append(pointTechTable.records[p]['CODE'])
    pointFonc.append(pointTechTable.records[p]['TYPE_FONC'])
    pointStruc.append(pointTechTable.records[p]['TYPE_STRUC'])
    pointPrp.append(pointTechTable.records[p]['PROPRIETAI'])
    rempli.append(pointTechTable.records[p]['REMPLA_APP'])
# FROM THE FCI
fciNom = []
fciCode = []
for f in range(0, fcilen):
    try:
        fciNom.append(fciTable.records[f]['NOM'])
    except KeyError:
        fciNom.append(fciTable.records[f]['N__'])

    fciCode.append(fciTable.records[f]['FCI'])

allcable = []
cablePoint = []
typeFonc = []
typeStruc = []
cablePTProp = []
capacity = []
for k in range(0, joinlen):
    allcable.append(joinTable.records[k]['NOM'])
    cablePoint.append(joinTable.records[k]['NOM_2'])
    typeFonc.append(joinTable.records[k]['TYPE_FONC_'])
    capacity.append(joinTable.records[k]['CAPACITE'])
    typeStruc.append(joinTable.records[k]['TYPE_STR_1'])
    cablePTProp.append(joinTable.records[k]['PROPRIET_1'])
# ###################### define  the base header ##############################
sheet = xlsxwriter.worksheet.Worksheet
listErour = []


# ################# normale ############
def baseHeader(w: sheet):
    w.write('A1', "Ligne 1", header)
    w.write('B1', 'Ligne 2', header)
    w.write('C1', 'Ligne 3', header)
    w.write('D1', 'LIGNE 4', header)
    w.write('E1', 'LIGNE 5', header)
    w.write('F1', 'LIGNE 6', header)
    w.merge_range('G1:G2', 'Qté', header)
    w.write('A2', "Projet", header)
    w.write('B2', 'Capacité', header)
    w.write('C2', 'Boite / Câble', header)
    w.write('D2', 'Date', header)
    w.write('E2', 'N°FCI', header)
    w.write('F2', '', header)


# #################################### function  ########################
def getPointCode(boite):
    index = boiteCode.index(boite)
    idPrent = boiteIdParent[index]
    return idPrent


# def getCapacity(cable):
#     index = cableName.index(cable)
#     cap = cableCapacity[index]
#     return cap


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
    return pointTech, idPrent


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


def duplicates(lst, item):
    return [i for i, x in enumerate(lst) if x == item]


def fillInAllTable(cable, pointCode, index):
    allcable.append(cable)
    cablePoint.append(pointCode)
    typeFonc.append(pointFonc[index])
    typeStruc.append(pointStruc[index])
    cablePTProp.append(pointPrp[index])


def getCapacity(point):
    index = cablePoint.index(point)
    cap = capacity[index]
    return int(cap)


def getBoite(code):
    try:
        index = boiteIdParent.index(code)
        return boiteCode[index]
    except:
        index = ""
        return index


def getNumAffaire(point):
    try:
        index = namEned.index(point)
        return codeEned[index]
    except:
        index = "CODE AFFAIRE non"
        return index


vonder = "VENU"


def boiteEtiqueteFill(boites, k, totale: sheet):
    for b in boites:

        if not str(b).startswith('SRO'):
            pointTech, idParent = getPointTech(b)
            index = getpointIndex(b)
            prop = str(pointPrp[index])
            if prop.startswith("O"):
                fcicode = getFci(pointTech)
                if fcicode is not None:
                    totale.write('A' + str(k), vonder, border)
                    totale.write('B' + str(k), '', border)
                    totale.write('C' + str(k), b, border)
                    totale.write('D' + str(k), "", border)
                    totale.write('E' + str(k), fcicode, border)
                    totale.write('F' + str(k), "", border)
                    totale.write('G' + str(k), '', border)
                    totale.write('H' + str(k), '', yellow)
                    k += 1
                else:
                    totale.write('A' + str(k), vonder, border)
                    totale.write('B' + str(k), '', border)
                    totale.write('C' + str(k), b, border)
                    totale.write('D' + str(k), "", border)
                    totale.write('E' + str(k), sro, border)
                    totale.write('F' + str(k), "", border)
                    totale.write('G' + str(k), '', border)
                    totale.write('H' + str(k), '', yellow)
                    k += 1
            else:
                # ############################
                totale.write('A' + str(k), vonder, border)
                totale.write('B' + str(k), '', border)
                totale.write('C' + str(k), b, border)
                totale.write('D' + str(k), "", border)
                totale.write('E' + str(k), idParent, border)
                totale.write('F' + str(k), "", border)
                totale.write('G' + str(k), '', border)
                totale.write('H' + str(k), '', yellow)
                k += 1

    return k


def cableEtiqueteFill(cables, totale: sheet):
    max = len(cables)
    k = 3
    for i in range(0, max):
        cable = cables[i]
        cap = int(capacity[i])
        point = cablePoint[i]
        prop = cablePTProp[i]
        if prop.startswith("OR"):
            fcicode = getFci(point)
            totale.write('A' + str(k), vonder, border)
            totale.write('B' + str(k), str(cap) + "Fo", border)
            totale.write('C' + str(k), cable, border)
            totale.write('D' + str(k), date, border)
            if fcicode is  None:
                fcicode = sro
            totale.write('E' + str(k), fcicode, border)
            totale.write('F' + str(k), "", border)
            totale.write('G' + str(k), '', border)
            totale.write('H' + str(k), '', yellow)
            k += 1
        elif prop.startswith("VENDEE"):
            totale.write('A' + str(k), vonder, border)
            totale.write('B' + str(k), str(cap) + "Fo", border)
            totale.write('C' + str(k), cable, border)
            totale.write('D' + str(k), date, border)
            totale.write('E' + str(k), sro, border)
            totale.write('F' + str(k), "", border)
            totale.write('G' + str(k), '', border)
            totale.write('H' + str(k), '', yellow)
            k += 1
        else:
            numAffaire = getNumAffaire(point)
            totale.write('A' + str(k), vonder, border)
            totale.write('B' + str(k), str(cap) + "Fo", border)
            totale.write('C' + str(k), cable, border)
            totale.write('D' + str(k), date, border)
            totale.write('E' + str(k), numAffaire, border)
            totale.write('F' + str(k), "", border)
            totale.write('G' + str(k), '', border)
            totale.write('H' + str(k), '', yellow)
            k += 1
    return k


def appuiEtiqueteFill(points, k, totale: sheet):
    for i in range(0, len(points)):
        point = points[i]
        code = pointCode[i]
        boite = getBoite(code)
        prop = pointPrp[i]
        typeS = str(pointStruc[i])
        if typeS.startswith("APPUI"):
            if prop.startswith("OR"):
                romp = str(rempli[i])
                fci = getFci(point)
                totale.write('A' + str(k), vonder, border)
                try:
                    cap = getCapacity(point)
                    totale.write('B' + str(k), str(cap) + "Fo", border)
                except:
                    cap = ""
                    totale.write('B' + str(k), cap, border)
                    print(point)
                totale.write('C' + str(k), boite, border)
                totale.write('D' + str(k), "", border)
                totale.write('E' + str(k), fci, border)
                totale.write('F' + str(k), "", border)
                totale.write('G' + str(k), '', border)
                totale.write('H' + str(k), '', green)
                k += 1
                if romp.startswith("OUI") or romp.startswith("oui"):
                    totale.write('A' + str(k), vonder, border)
                    totale.write('B' + str(k), "", border)
                    totale.write('C' + str(k)," ", border)
                    totale.write('D' + str(k), "", border)
                    totale.write('E' + str(k), point, border)
                    totale.write('F' + str(k), "", border)
                    totale.write('G' + str(k), '', border)
                    totale.write('H' + str(k), '', bleu)
                    k += 1
            elif prop.startswith("VENDEE"):
                totale.write('A' + str(k), vonder, border)
                totale.write('B' + str(k), code, border)
                totale.write('C' + str(k), boite, border)
                totale.write('D' + str(k), date, border)
                if str(point).startswith("P"):
                    point = point[4:]
                totale.write('E' + str(k), point[0:5], border)
                totale.write('F' + str(k), "", border)
                totale.write('G' + str(k), '', border)
                totale.write('H' + str(k), '', white)
                k += 1


baseHeader(totaleSheet)
k = cableEtiqueteFill(allcable, totaleSheet)
k = boiteEtiqueteFill(boiteCode, k, totaleSheet)
appuiEtiqueteFill(pointNom, k, totaleSheet)
workbook.close()
