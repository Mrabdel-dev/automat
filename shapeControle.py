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
sheet = xlsxwriter.worksheet.Worksheet
# FROM THE BOITE OPTIQUE
boiteCode = []  # name of the boite
boiteAmount = []  # name of the
boiteIdParent = []  # AMOUNT CABLE
for j in range(0, boiteLen):
    boiteCode.append(boiteTable.records[j]['CODE'])
    boiteIdParent.append(boiteTable.records[j]['ID_PARENT'])
    boiteAmount.append(boiteTable.records[j]['AMONT'])
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
suppInsee = []
for s in range(0, supplen):
    suppAmount.append(supportTable.records[s]['AMONT'])
    suppAval.append(supportTable.records[s]['AVAL'])
    suppExi.append(supportTable.records[s]['NOM'])
    suppInsee.append(supportTable.records[s]['INSEE'])
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
workbook = xlsxwriter.Workbook('shapeControleResulte' + SRO + '.xlsx')
erourSheet = workbook.add_worksheet("allEroorResulte")
bold = workbook.add_format({'bold': True, "border": 1, 'bg_color': '#B7B5B5'})
bold.set_center_across()
bold1 = workbook.add_format({'bold': True, "border": 1})
bold1.set_center_across()
header = workbook.add_format({'bold': True, 'border': 1, 'bg_color': 'red'})
header1 = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#E17070'})
header2 = workbook.add_format({'bold': True, 'border': 1, 'bg_color': 'green'})


# ####################################################################################

def baseHeader():
    erourSheet.merge_range('A1:D1', "supportTableEroor", bold)
    erourSheet.write("A2", "suppExi", bold1)
    erourSheet.write("B2", "Insee", bold1)
    erourSheet.write("C2", "Amont", bold1)
    erourSheet.write("D2", "Aval", bold1)
    erourSheet.merge_range('G1:H1', "CableTableEroor", bold)
    erourSheet.write("G2", "CABLE", bold1)
    erourSheet.write("H2", "BOITE", bold1)
    erourSheet.merge_range('K1:M1', "BoiteTableEroor", bold)
    erourSheet.write("K2", "BOITE", bold1)
    erourSheet.write("L2", "CABLE", bold1)
    erourSheet.write("M2", "ID_Parent", bold1)


def suppTest(lin: int):
    for i in range(0, len(suppExi)):
        ens = str(suppInsee[i])
        aval = str(suppAval[i])
        amont = str(suppAmount[i])
        if ens in aval or ens in amont:
            print("yesssssssssss")
        else:
            if amont.startswith("C") or amont.startswith("P") or aval.startswith("C") or aval.startswith("P"):
                erourSheet.write("A" + str(lin), suppExi[i], header2)
                erourSheet.write("B" + str(lin), ens, header2)
                erourSheet.write("C" + str(lin), amont, header2)
                erourSheet.write("D" + str(lin), aval, header2)
                lin += 1
            else:
                erourSheet.write("A" + str(lin), suppExi[i], header)
                erourSheet.write("B" + str(lin), ens, header)
                erourSheet.write("C" + str(lin), amont, header)
                erourSheet.write("D" + str(lin), aval, header)
                lin += 1


def cableTest(lin: int):
    for c in range(0, len(cableName)):
        cable = str(cableName[c])
        boite = str(cableExtremity[c])
        if cable[3:] == boite[3:]:
            print("yes")
        else:
            erourSheet.write("G" + str(lin), cable, header)
            erourSheet.write("H" + str(lin), boite, header)
            lin += 1


def boiteTest(lin: int):
    for b in range(0, len(boiteTable)):
        boite = str(boiteCode[b])
        idPrent = str(boiteIdParent[i])
        cable = str(boiteAmount[b])
        if boite[3:] != cable[3:]:
            erourSheet.write("K" + str(lin), boite, header)
            erourSheet.write("L" + str(lin), cable, header)
            erourSheet.write("M" + str(lin), idPrent, header)
            lin += 1
        try:
            try:
                index = suppAmount.index(idPrent)
                print(index, "amount")
            except:
                index = suppAval.index(idPrent)
                print(index, "aval")
        except:
            erourSheet.write("K" + str(lin), boite, header1)
            erourSheet.write("L" + str(lin), cable, header1)
            erourSheet.write("M" + str(lin), idPrent, header1)
            lin += 1


baseHeader()
suppTest(3)
cableTest(3)
boiteTest(3)
workbook.close()
