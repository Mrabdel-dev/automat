from dbfread import DBF, FieldParser
import xlsxwriter
from os import walk

workbook = xlsxwriter.Workbook('statistique_-21-017.xlsx')
header = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#C4E5F7'})
bold = workbook.add_format({'bold': True, "border": 1})
w = workbook.add_worksheet('statistique')
w.write("A1", "SRO-cable", header)
w.write("B1", "totaleConduite", header)
w.write("C1", "totaleAerien", header)
w.write("E1", "SRO-SUPP", header)
w.write("F1", "totale", header)
w.write("H1", "SRO-cableTrans", header)
w.write("I1", "totaleConduiteT", header)
w.write("J1", "totaleAerienT", header)
w.write("L1", "SRO-SUPpTrans", header)
w.write("M1", "totaleTrans", header)
monRepertoire = 'ALL CABLE/'
monRepertoire2 = 'ALL SUPPORT/'
monRepertoire3 = 'ALL CABLET/'
monRepertoire4 = 'ALL SUPPORTT/'
listeFichiersCable = []
for (repertoire, sousRepertoires, fichiers) in walk(monRepertoire):
    listeFichiersCable.extend(fichiers)

listeFichiersSupp = []
for (repertoire, sousRepertoires, fichiers) in walk(monRepertoire2):
    listeFichiersSupp.extend(fichiers)
listeFichiersCableT = []
for (repertoire, sousRepertoires, fichiers) in walk(monRepertoire3):
    listeFichiersCableT.extend(fichiers)

listeFichiersSuppT = []
for (repertoire, sousRepertoires, fichiers) in walk(monRepertoire4):
    listeFichiersSuppT.extend(fichiers)
SRO = ""
i = 0
f = 2
print(listeFichiersCable)
for d in listeFichiersCable:
    SRO = str(d)[0:11]
    try:
        cableTable = DBF(monRepertoire + str(d), load=True, encoding='iso-8859-1')
        filedCableNam = cableTable.field_names
        cableName = []  # NAME OF THE CABLE
        cableOrigin = []  # WHERE THEY COME FROM
        cableType = []  # WHERE HE GO IN
        cableLong = []  # CAPACITY OF THE CABLE
        cableLen = len(cableTable)
        for i in range(0, cableLen):
            cableName.append(cableTable.records[i]['NOM'])
            cableType.append(cableTable.records[i]['TYPE_STRUC'])
            cableLong.append(cableTable.records[i]['LGR_REELLE'])
        totaleConduite = 0.0
        totaleAerien = 0.0
        for c, t, l in zip(cableName, cableType, cableLong):
            if str(t) == "AERIEN":
                try:
                    totaleAerien += float(l)
                except TypeError:
                    totaleAerien += 0.0
            elif str(t) == "CONDUITE":
                try:
                    totaleConduite += float(l)
                except TypeError:
                    totaleConduite += 0.0
            else:
                continue
        w.write("A" + str(f), SRO, bold)
        w.write("B" + str(f), totaleConduite, bold)
        w.write("C" + str(f), totaleAerien, bold)
        f += 1
    except ValueError:
        print(d)
        continue

# FROM THE SUPPORT
k = 0
t = 2
while k < len(listeFichiersSupp):
    SRO = str(listeFichiersSupp[k])[0:11]
    supportTable = DBF(monRepertoire2 + listeFichiersSupp[k], load=True, encoding='iso-8859-1')
    suppCode = []
    suppGest = []
    suppUtil = []
    suppLong = []

    supplen = len(supportTable)
    for s in range(0, supplen):
        suppCode.append(supportTable.records[s]['CODE'])
        suppGest.append(supportTable.records[s]['GESTIONNAI'])
        suppLong.append(supportTable.records[s]['LG_REELLE'])
        suppUtil.append(supportTable.records[s]['UTILISATIO'])

    totaleSupprt = 0.0
    for c, g, l, u in zip(suppCode, suppGest,suppLong, suppUtil):
        if str(c).startswith("TRA"):
            if str(g).startswith("ALT") and str(u) == "D":
                try:
                    totaleSupprt += float(l)
                except TypeError:
                    totaleSupprt += 0.0
    w.write("E" + str(t), SRO, bold)
    w.write("F" + str(t), totaleSupprt, bold)
    t += 1
    k += 1
# ######################## transport #############################################
i = 0
f = 2
for d in listeFichiersCableT:
    SRO = str(d)[0:11]
    try:
        cableTable = DBF(monRepertoire3 + str(d), load=True, encoding='iso-8859-1')
        filedCableNam = cableTable.field_names
        cableName = []  # NAME OF THE CABLE
        cableOrigin = []  # WHERE THEY COME FROM
        cableType = []  # WHERE HE GO IN
        cableLong = []  # CAPACITY OF THE CABLE
        cableLen = len(cableTable)
        for i in range(0, cableLen):
            cableName.append(cableTable.records[i]['NOM'])
            cableType.append(cableTable.records[i]['TYPE_STRUC'])
            cableLong.append(cableTable.records[i]['LGR_REELLE'])
        totaleConduite = 0.0
        totaleAerien = 0.0
        for c, t, l in zip(cableName, cableType, cableLong):
            if str(t) == "AERIEN":
                try:
                    totaleAerien += float(l)
                except TypeError:
                    totaleAerien += 0.0
            elif str(t) == "CONDUITE":
                try:
                    totaleConduite += float(l)
                except TypeError:
                    totaleConduite += 0.0
            else:
                continue
        w.write("H" + str(f), SRO, bold)
        w.write("I" + str(f), totaleConduite, bold)
        w.write("J" + str(f), totaleAerien, bold)
        f += 1
    except ValueError:
        print(d)
        continue

# FROM THE SUPPORT
k = 0
t = 2
while k < len(listeFichiersSuppT):
    SRO = str(listeFichiersSuppT[k])[0:11]
    supportTable = DBF(monRepertoire4 + listeFichiersSuppT[k], load=True, encoding='iso-8859-1')
    suppCode = []
    suppGest = []
    suppUtil = []
    suppLong = []

    supplen = len(supportTable)
    for s in range(0, supplen):
        suppCode.append(supportTable.records[s]['CODE'])
        suppGest.append(supportTable.records[s]['GESTIONNAI'])
        suppLong.append(supportTable.records[s]['LG_REELLE'])
        suppUtil.append(supportTable.records[s]['UTILISATIO'])

    totaleSupprt = 0.0
    for c, g, l, u in zip(suppCode, suppGest,suppLong, suppUtil):
        if str(c).startswith("TRA"):
            if str(g).startswith("ALT") and str(u) != "D":
                try:
                    totaleSupprt += float(l)
                except TypeError:
                    totaleSupprt += 0.0
    w.write("L" + str(t), SRO, bold)
    w.write("M" + str(t), totaleSupprt, bold)
    t += 1
    k += 1
workbook.close()
