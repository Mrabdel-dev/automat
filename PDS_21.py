import xlrd
import sys
import xlsxwriter
import pandas as pd
import bisect
import datetime
import qgis_plugin_repo

# import win32com.client as win32

path_to_boite_optique = "C:/Users/BE24/Desktop/Projet Fibre 21/SID/21_011_069_EXE_CCE_SIO_MAN_001_B_AI/21_011_069_BOITE_OPTIQUE_B.shx"
blayer = QgsVectorLayer(path_to_boite_optique, "boite_optique", "ogr")

path_to_CABLE_OPTIQUE = "C:/Users/BE24/Desktop/Projet Fibre 21/SID/21_011_069_EXE_CCE_SIO_MAN_001_B_AI/21_011_069_CABLE_OPTIQUE_B.shx"
clayer = QgsVectorLayer(path_to_CABLE_OPTIQUE, "CABLE_OPTIQUE", "ogr")

boite = [];
entree = [];
interco = [];
reference = [];
nbf = []
cable = [];
c_origine = [];
c_extremite = [];
capacite = []
for feature in blayer.getFeatures():
    name = feature["NOM"]
    amont = feature["AMONT"]
    ext = feature["INTERCO"]
    ref = feature["REFERENCE"]
    fu = feature["NBFUTILE"]
    nbf.append(fu)
    reference.append(ref)
    boite.append(name)
    entree.append(amont)
    interco.append(ext)
for feature in clayer.getFeatures():
    name = feature["NOM"]
    cap = feature["CAPACITE"]
    origine = feature["ORIGINE"]
    extrem = feature["EXTREMITE"]
    cable.append(name)
    capacite.append(cap)
    c_origine.append(origine)
    c_extremite.append(extrem)


def duplicates(lst, item):
    return [i for i, x in enumerate(lst) if x == item]


workbook = xlsxwriter.Workbook('C:\\Users\\BE24\\Desktop\\Projet Fibre 21\\PDS\\21_011_069.xlsx')

bold = workbook.add_format({'bold': True, "border": 1})
border = workbook.add_format({"border": 1})
retour = workbook.add_format({"bg_color": '#CD5C5C', "border": 1})
cassette = workbook.add_format({"bg_color": '#A9A9A9', "border": 1})
cell_format0 = workbook.add_format({"bg_color": '#E6E6FA', "border": 1})
cell_format = workbook.add_format({"bg_color": 'red', "border": 1})
cell_format1 = workbook.add_format({"bg_color": 'blue', "border": 1})
cell_format2 = workbook.add_format({"bg_color": '#00FF00', "border": 1})
cell_format3 = workbook.add_format({"bg_color": 'yellow', "border": 1})
cell_format4 = workbook.add_format({"bg_color": '#BF00FF', "border": 1})
cell_format5 = workbook.add_format({"bg_color": 'white', "border": 1})
cell_format6 = workbook.add_format({"bg_color": '#FFBF00', "border": 1})
cell_format7 = workbook.add_format({"bg_color": '#828282', "border": 1})
cell_format8 = workbook.add_format({"bg_color": '#816B56', "border": 1})
cell_format9 = workbook.add_format({"bg_color": '#333333', "border": 1})
cell_format10 = workbook.add_format({"bg_color": '#00FFBF', "border": 1})
cell_format11 = workbook.add_format({"bg_color": '#FFAAD4', "border": 1})

now = datetime.datetime.now()
date = now.strftime("%d/%m/%Y")

T = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31,
     32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48]
N = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12];
locs = []
t = 0;
j = 0;
i = 12;
count = 0
l = len(boite)
m = 12 + l
origi = []
start_at = -1

while j < l:
    w = workbook.add_worksheet(c_extremite[j])
    w.write('A7', 'Etiquette : ' + c_extremite[j], border)
    w.write('O3', date, bold)
    w.write('N3', 'Date de modification : ', bold)
    w.write('A11', 'Entrée', bold)
    w.write('B11', 'Id', bold)
    w.write('C11', 'Capacité', bold)
    w.write('D11', 'N°         ', bold)
    w.write('E11', 'N° Tube', bold)
    w.write('F11', 'N° Fibre', bold)
    w.write('G11', 'Cassette', bold)
    w.write('H11', 'Etat fibre', bold)
    w.write('I11', 'N° Fibre', bold)
    w.write('J11', 'N° Tube', bold)
    w.write('K11', 'N°       ', bold)
    w.write('L11', 'Capacité', bold)
    w.write('M11', '', bold)
    w.write('N11', 'Sortie', bold)
    w.write('O11', 'Statut', bold)
    w.write('P11', 'Client', bold)
    d = boite.index(c_extremite[j])
    w.write('A9', 'Reference : ' + reference[d], border)
    w.write('A1', '<- RETOUR:', retour)

    if interco[d] != "EXTREMITE":
        # orig = c_origine.index(c_extremite[j])
        val = c_extremite[j]
        orig = c_origine.index(val)
        origi = duplicates(c_origine, val)
        m = 1
        for k in origi:
            w.write('Q' + str(m), c_extremite[k], border)
            fa = boite.index(c_extremite[k])
            # w.write('R'+str(m),nbf[fa],border)
            m += 1
            # print(c_origine[k],c_extremite[k])
        # w.write('Q1',c_extremite[orig],border)
        w.write('P1', 'SUIVANT ->:', retour)
        w.write('A2', c_origine[j], border)

    elif interco[d] == "EXTREMITE":
        w.write('A2', c_origine[j], border)
        w.write('Q1', 'EXTREMITE', retour)
    c = int(capacite[j])
    f = int(c + 12)
    for i in range(12, f):
        w.write('A' + str(i), cable[j], border)
        w.write('B' + str(i), '', cell_format0)
        w.write('C' + str(i), c, cell_format0)
        w.write('G' + str(i), 1, cassette)
        w.write('H12', 'LIBRE', border)
        w.write('H14', 'STOCKEE', border)
        w.write('H16', 'EN PASSAGE', border)
        w.write('H19', 'EN ATTENTE', border)
        w.write('H23', 'EPISSUREE', border)
        w.write('I' + str(i), '', border)
        w.write('J' + str(i), '', border)
        w.write('K' + str(i), '', border)
        if interco[d] != "EXTREMITE":
            ma = 12
            origi = duplicates(c_origine, c_extremite[j])
            for k in origi:
                w.write('L' + str(ma), capacite[k], border)
                w.write('N' + str(ma), cable[k], border)
                ma += 2
            # w.write('L'+ str(i), capacite[orig], border)
            w.write('M' + str(i), '', border)
        elif interco[d] == "EXTREMITE":
            w.write('L' + str(i), '', border)
            w.write('M' + str(i), '', border)
            w.write('N' + str(i), '', border)

        w.write('O' + str(i), '', border)
        w.write('P' + str(i), '', border)

        if i < 24:
            w.write_column('D12', N, border)
            w.write('E' + str(i), T[0], cell_format)
            # w.write_column('F12', N, border)
            if i == 12:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 13:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 14:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 15:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 16:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 17:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 18:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 19:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 20:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 21:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 22:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 23:
                w.write('F' + str(i), N[11], cell_format11)

        elif i > 23 and i < 36:
            w.write_column('D24', N, border)
            w.write('E' + str(i), T[1], cell_format1)
            # w.write_column('F24', N, border)
            if i == 24:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 25:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 26:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 27:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 28:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 29:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 30:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 31:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 32:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 33:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 34:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 35:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 35 and i < 48:
            w.write_column('D36', N, border)
            w.write('E' + str(i), T[2], cell_format2)
            # w.write_column('F36', N, border)
            if i == 36:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 37:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 38:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 39:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 40:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 41:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 42:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 43:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 44:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 45:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 46:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 47:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 47 and i < 60:
            w.write_column('D48', N, border)
            w.write('E' + str(i), T[3], cell_format3)
            # w.write_column('F48', N, border)
            if i == 48:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 49:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 50:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 51:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 52:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 53:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 54:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 55:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 56:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 57:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 58:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 59:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 59 and i < 72:
            w.write_column('D60', N, border)
            w.write('E' + str(i), T[4], cell_format4)
            # w.write_column('F60', N, border)
            if i == 60:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 61:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 62:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 63:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 64:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 65:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 66:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 67:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 68:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 69:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 70:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 71:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 71 and i < 84:
            w.write_column('D72', N, border)
            w.write('E' + str(i), T[5], cell_format5)
            # w.write_column('F72', N, border)
            if i == 72:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 73:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 74:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 75:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 76:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 77:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 78:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 79:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 80:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 81:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 82:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 83:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 83 and i < 96:
            w.write_column('D84', N, border)
            w.write('E' + str(i), T[6], cell_format6)
            # w.write_column('F84', N, border)
            if i == 84:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 85:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 86:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 87:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 88:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 89:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 90:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 91:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 92:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 93:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 94:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 95:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 95 and i < 108:
            w.write_column('D96', N, border)
            w.write('E' + str(i), T[7], cell_format7)
            # w.write_column('F96', N, border)
            if i == 96:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 97:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 98:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 99:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 100:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 101:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 102:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 103:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 104:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 105:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 106:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 107:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 107 and i < 120:
            w.write_column('D108', N, border)
            w.write('E' + str(i), T[8], cell_format8)
            # w.write_column('F108', N, border)
            if i == 108:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 109:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 110:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 111:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 112:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 113:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 114:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 115:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 116:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 117:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 118:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 119:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 119 and i < 132:
            w.write_column('D120', N, border)
            w.write('E' + str(i), T[9], cell_format9)
            # w.write_column('F120', N, border)
            if i == 120:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 121:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 122:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 123:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 124:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 125:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 126:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 127:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 128:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 129:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 130:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 131:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 131 and i < 144:
            w.write_column('D132', N, border)
            w.write('E' + str(i), T[10], cell_format10)
            # w.write_column('F132', N, border)
            if i == 132:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 133:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 134:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 135:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 136:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 137:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 138:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 139:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 140:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 141:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 142:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 143:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 143 and i < 156:
            w.write_column('D144', N, border)
            w.write('E' + str(i), T[11], cell_format11)
            # w.write_column('F144', N, border)
            if i == 144:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 145:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 146:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 147:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 148:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 149:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 150:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 151:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 152:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 153:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 154:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 155:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 155 and i < 168:
            w.write_column('D156', N, border)
            w.write('E' + str(i), T[12], cell_format)
            # w.write_column('F156', N, border)
            if i == 156:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 157:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 158:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 159:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 160:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 161:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 162:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 163:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 164:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 165:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 166:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 167:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 167 and i < 180:
            w.write_column('D168', N, border)
            w.write('E' + str(i), T[13], cell_format1)
            # w.write_column('F168', N, border)
            if i == 168:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 169:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 170:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 171:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 172:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 173:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 174:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 175:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 176:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 177:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 178:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 179:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 179 and i < 192:
            w.write_column('D180', N, border)
            w.write('E' + str(i), T[14], cell_format2)
            # w.write_column('F180', N, border)
            if i == 180:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 181:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 182:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 183:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 184:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 185:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 186:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 187:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 188:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 189:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 190:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 191:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 191 and i < 204:
            w.write_column('D192', N, border)
            w.write('E' + str(i), T[15], cell_format3)
            # w.write_column('F192', N, border)
            if i == 192:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 193:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 194:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 195:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 196:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 197:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 198:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 199:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 200:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 201:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 202:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 203:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 203 and i < 216:
            w.write_column('D204', N, border)
            w.write('E' + str(i), T[16], cell_format4)
            # w.write_column('F204', N, border)
            if i == 204:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 205:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 206:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 207:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 208:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 209:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 210:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 211:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 212:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 213:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 214:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 215:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 215 and i < 228:
            w.write_column('D216', N, border)
            w.write('E' + str(i), T[17], cell_format5)
            # w.write_column('F216', N, border)
            if i == 216:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 217:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 218:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 219:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 220:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 221:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 222:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 223:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 224:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 225:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 226:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 227:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 227 and i < 240:
            w.write_column('D228', N, border)
            w.write('E' + str(i), T[18], cell_format6)
            # w.write_column('F228', N, border)
            if i == 228:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 229:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 230:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 231:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 232:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 233:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 234:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 235:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 236:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 237:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 238:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 239:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 239 and i < 252:
            w.write_column('D240', N, border)
            w.write('E' + str(i), T[19], cell_format7)
            # w.write_column('F240', N, border)
            if i == 240:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 241:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 242:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 243:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 244:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 245:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 246:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 247:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 248:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 249:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 250:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 251:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 251 and i < 264:
            w.write_column('D252', N, border)
            w.write('E' + str(i), T[20], cell_format8)
            # w.write_column('F252', N, border)
            if i == 252:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 253:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 254:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 255:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 256:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 257:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 258:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 259:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 260:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 261:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 262:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 263:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 263 and i < 276:
            w.write_column('D264', N, border)
            w.write('E' + str(i), T[21], cell_format9)
            # w.write_column('F264', N, border)
            if i == 264:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 265:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 266:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 267:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 268:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 269:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 270:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 271:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 272:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 273:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 274:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 275:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 275 and i < 288:
            w.write_column('D276', N, border)
            w.write('E' + str(i), T[22], cell_format10)
            # w.write_column('F276', N, border)
            if i == 276:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 277:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 278:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 279:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 280:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 281:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 282:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 283:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 284:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 285:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 286:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 287:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 287 and i < 300:
            w.write_column('D288', N, border)
            w.write('E' + str(i), T[23], cell_format11)
            # w.write_column('F288', N, border)
            if i == 288:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 289:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 290:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 291:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 292:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 293:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 294:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 295:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 296:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 297:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 298:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 299:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 299 and i < 312:
            w.write_column('D300', N, border)
            w.write('E' + str(i), T[24], cell_format)
            if i == 300:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 301:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 302:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 303:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 304:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 305:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 306:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 307:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 308:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 309:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 310:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 311:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 311 and i < 324:
            w.write_column('D312', N, border)
            w.write('E' + str(i), T[25], cell_format1)
            if i == 312:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 313:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 314:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 315:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 316:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 317:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 318:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 319:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 320:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 321:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 322:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 323:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 323 and i < 336:
            w.write_column('D324', N, border)
            w.write('E' + str(i), T[26], cell_format2)
            if i == 324:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 325:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 326:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 327:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 328:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 329:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 330:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 331:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 332:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 333:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 334:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 335:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 335 and i < 348:
            w.write_column('D336', N, border)
            w.write('E' + str(i), T[27], cell_format3)
            if i == 336:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 337:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 338:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 339:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 340:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 341:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 342:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 343:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 344:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 345:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 346:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 347:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 347 and i < 360:
            w.write_column('D348', N, border)
            w.write('E' + str(i), T[28], cell_format4)
            if i == 348:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 349:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 350:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 351:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 352:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 353:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 354:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 355:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 356:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 357:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 358:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 359:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 359 and i < 372:
            w.write_column('D360', N, border)
            w.write('E' + str(i), T[29], cell_format5)
            if i == 360:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 361:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 362:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 363:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 364:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 365:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 366:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 367:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 368:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 369:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 370:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 371:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 371 and i < 384:
            w.write_column('D372', N, border)
            w.write('E' + str(i), T[30], cell_format6)
            if i == 372:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 373:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 374:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 375:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 376:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 377:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 378:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 379:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 380:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 381:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 382:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 383:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 383 and i < 396:
            w.write_column('D384', N, border)
            w.write('E' + str(i), T[31], cell_format7)
            if i == 384:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 385:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 386:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 387:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 388:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 389:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 390:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 391:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 392:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 393:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 394:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 395:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 395 and i < 408:
            w.write_column('D396', N, border)
            w.write('E' + str(i), T[32], cell_format8)
            if i == 396:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 397:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 398:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 399:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 400:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 401:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 402:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 403:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 404:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 405:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 406:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 407:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 407 and i < 420:
            w.write_column('D408', N, border)
            w.write('E' + str(i), T[33], cell_format9)
            if i == 408:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 409:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 410:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 411:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 412:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 413:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 414:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 415:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 416:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 417:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 418:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 419:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 419 and i < 432:
            w.write_column('D420', N, border)
            w.write('E' + str(i), T[34], cell_format10)
            if i == 420:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 421:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 422:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 423:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 424:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 425:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 426:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 427:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 428:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 429:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 430:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 431:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 431 and i < 444:
            w.write_column('D432', N, border)
            w.write('E' + str(i), T[35], cell_format11)
            if i == 432:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 433:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 434:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 435:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 436:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 437:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 438:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 439:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 440:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 441:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 442:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 443:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 443 and i < 456:
            w.write_column('D444', N, border)
            w.write('E' + str(i), T[36], cell_format)
            if i == 444:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 445:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 446:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 447:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 448:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 449:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 450:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 451:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 452:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 453:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 454:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 455:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 455 and i < 468:
            w.write_column('D456', N, border)
            w.write('E' + str(i), T[37], cell_format1)
            if i == 456:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 457:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 458:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 459:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 460:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 461:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 462:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 463:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 464:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 465:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 466:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 467:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 467 and i < 480:
            w.write_column('D468', N, border)
            w.write('E' + str(i), T[38], cell_format2)
            if i == 458:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 469:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 470:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 471:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 472:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 473:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 474:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 475:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 476:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 477:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 478:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 479:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 479 and i < 492:
            w.write_column('D480', N, border)
            w.write('E' + str(i), T[39], cell_format3)
            if i == 480:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 481:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 482:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 483:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 484:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 485:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 486:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 487:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 488:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 489:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 490:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 491:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 491 and i < 504:
            w.write_column('D492', N, border)
            w.write('E' + str(i), T[40], cell_format4)
            if i == 492:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 493:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 494:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 495:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 496:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 497:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 498:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 499:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 500:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 501:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 502:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 503:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 503 and i < 516:
            w.write_column('D504', N, border)
            w.write('E' + str(i), T[41], cell_format5)
            if i == 504:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 505:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 506:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 507:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 508:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 509:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 510:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 511:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 512:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 513:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 514:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 515:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 515 and i < 528:
            w.write_column('D516', N, border)
            w.write('E' + str(i), T[42], cell_format6)
            if i == 516:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 517:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 518:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 519:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 520:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 521:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 522:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 523:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 524:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 525:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 526:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 527:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 527 and i < 540:
            w.write_column('D528', N, border)
            w.write('E' + str(i), T[43], cell_format7)
            if i == 528:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 529:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 530:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 531:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 532:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 533:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 534:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 535:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 536:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 537:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 538:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 539:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 539 and i < 552:
            w.write_column('D540', N, border)
            w.write('E' + str(i), T[44], cell_format8)
            if i == 540:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 541:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 542:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 543:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 544:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 545:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 546:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 547:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 548:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 549:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 550:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 551:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 551 and i < 564:
            w.write_column('D552', N, border)
            w.write('E' + str(i), T[45], cell_format9)
            if i == 552:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 553:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 554:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 555:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 556:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 557:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 558:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 559:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 560:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 561:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 562:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 563:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 563 and i < 576:
            w.write_column('D564', N, border)
            w.write('E' + str(i), T[46], cell_format10)
            if i == 564:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 565:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 566:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 567:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 568:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 569:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 570:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 571:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 572:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 573:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 574:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 575:
                w.write('F' + str(i), N[11], cell_format11)
        elif i > 575 and i < 588:
            w.write_column('D576', N, border)
            w.write('E' + str(i), T[47], cell_format11)
            if i == 576:
                w.write('F' + str(i), N[0], cell_format)
            elif i == 577:
                w.write('F' + str(i), N[1], cell_format1)
            elif i == 578:
                w.write('F' + str(i), N[2], cell_format2)
            elif i == 579:
                w.write('F' + str(i), N[3], cell_format3)
            elif i == 580:
                w.write('F' + str(i), N[4], cell_format4)
            elif i == 581:
                w.write('F' + str(i), N[5], cell_format5)
            elif i == 582:
                w.write('F' + str(i), N[6], cell_format6)
            elif i == 583:
                w.write('F' + str(i), N[7], cell_format7)
            elif i == 584:
                w.write('F' + str(i), N[8], cell_format8)
            elif i == 585:
                w.write('F' + str(i), N[9], cell_format9)
            elif i == 586:
                w.write('F' + str(i), N[10], cell_format10)
            elif i == 587:
                w.write('F' + str(i), N[11], cell_format11)
        i += 1
    for k in range(12, f):
        origi = duplicates(c_origine, c_extremite[j])
        mo = 1
        for ka in origi:
            w.write('Q' + str(mo), c_extremite[ka], bold)
            mo += 1
        if interco[d] != "EXTREMITE" and capacite[orig] == capacite[j] and len(origi) < 2:
            w.write('L' + str(k), capacite[orig], bold)
            w.write('N' + str(k), cable[orig], border)
            if k < 24:
                w.write_column('K12', N, border)
                w.write('J' + str(k), T[0], cell_format)
                # w.write_column('F12', N, border)
                if k == 12:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 13:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 14:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 15:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 16:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 17:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 18:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 19:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 20:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 21:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 22:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 23:
                    w.write('I' + str(k), N[11], cell_format11)

            elif k > 23 and k < 36:
                w.write_column('K24', N, border)
                w.write('J' + str(k), T[1], cell_format1)
                # w.write_column('F24', N, border)
                if k == 24:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 25:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 26:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 27:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 28:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 29:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 30:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 31:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 32:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 33:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 34:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 35:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 35 and k < 48:
                w.write_column('K36', N, border)
                w.write('J' + str(k), T[2], cell_format2)
                # w.write_column('F36', N, border)
                if k == 36:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 37:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 38:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 39:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 40:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 41:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 42:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 43:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 44:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 45:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 46:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 47:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 47 and k < 60:
                w.write_column('K48', N, border)
                w.write('J' + str(k), T[3], cell_format3)
                # w.write_column('F48', N, border)
                if k == 48:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 49:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 50:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 51:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 52:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 53:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 54:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 55:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 56:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 57:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 58:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 59:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 59 and k < 72:
                w.write_column('K60', N, border)
                w.write('J' + str(k), T[4], cell_format4)
                # w.write_column('F60', N, border)
                if k == 60:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 61:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 62:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 63:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 64:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 65:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 66:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 67:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 68:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 69:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 70:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 71:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 71 and k < 84:
                w.write_column('K72', N, border)
                w.write('J' + str(k), T[5], cell_format5)
                # w.write_column('F72', N, border)
                if k == 72:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 73:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 74:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 75:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 76:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 77:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 78:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 79:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 80:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 81:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 82:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 83:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 83 and k < 96:
                w.write_column('K84', N, border)
                w.write('J' + str(k), T[6], cell_format6)
                # w.write_column('F84', N, border)
                if k == 84:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 85:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 86:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 87:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 88:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 89:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 90:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 91:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 92:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 93:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 94:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 95:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 95 and k < 108:
                w.write_column('K96', N, border)
                w.write('J' + str(k), T[7], cell_format7)
                # w.write_column('F96', N, border)
                if k == 96:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 97:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 98:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 99:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 100:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 101:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 102:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 103:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 104:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 105:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 106:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 107:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 107 and k < 120:
                w.write_column('K108', N, border)
                w.write('J' + str(k), T[8], cell_format8)
                # w.write_column('F108', N, border)
                if k == 108:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 109:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 110:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 111:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 112:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 113:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 114:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 115:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 116:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 117:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 118:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 119:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 119 and k < 132:
                w.write_column('K120', N, border)
                w.write('J' + str(k), T[9], cell_format9)
                # w.write_column('F120', N, border)
                if k == 120:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 121:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 122:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 123:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 124:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 125:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 126:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 127:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 128:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 129:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 130:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 131:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 131 and k < 144:
                w.write_column('K132', N, border)
                w.write('J' + str(k), T[10], cell_format10)
                # w.write_column('F132', N, border)
                if k == 132:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 133:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 134:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 135:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 136:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 137:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 138:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 139:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 140:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 141:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 142:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 143:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 143 and k < 156:
                w.write_column('K144', N, border)
                w.write('J' + str(k), T[11], cell_format11)
                # w.write_column('F144', N, border)
                if k == 144:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 145:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 146:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 147:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 148:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 149:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 150:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 151:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 152:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 153:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 154:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 155:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 155 and k < 168:
                w.write_column('K156', N, border)
                w.write('J' + str(k), T[12], cell_format)
                # w.write_column('F156', N, border)
                if k == 156:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 157:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 158:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 159:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 160:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 161:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 162:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 163:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 164:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 165:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 166:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 167:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 167 and k < 180:
                w.write_column('K168', N, border)
                w.write('J' + str(k), T[13], cell_format1)
                # w.write_column('F168', N, border)
                if k == 168:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 169:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 170:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 171:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 172:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 173:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 174:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 175:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 176:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 177:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 178:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 179:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 179 and k < 192:
                w.write_column('K180', N, border)
                w.write('J' + str(k), T[14], cell_format2)
                # w.write_column('F180', N, border)
                if k == 180:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 181:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 182:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 183:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 184:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 185:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 186:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 187:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 188:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 189:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 190:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 191:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 191 and k < 204:
                w.write_column('K192', N, border)
                w.write('J' + str(k), T[15], cell_format3)
                # w.write_column('F192', N, border)
                if k == 192:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 193:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 194:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 195:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 196:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 197:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 198:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 199:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 200:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 201:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 202:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 203:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 203 and k < 216:
                w.write_column('K204', N, border)
                w.write('J' + str(k), T[16], cell_format4)
                # w.write_column('F204', N, border)
                if k == 204:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 205:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 206:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 207:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 208:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 209:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 210:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 211:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 212:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 213:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 214:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 215:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 215 and k < 228:
                w.write_column('K216', N, border)
                w.write('J' + str(k), T[17], cell_format5)
                # w.write_column('F216', N, border)
                if k == 216:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 217:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 218:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 219:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 220:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 221:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 222:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 223:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 224:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 225:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 226:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 227:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 227 and k < 240:
                w.write_column('K228', N, border)
                w.write('J' + str(k), T[18], cell_format6)
                # w.write_column('F228', N, border)
                if k == 228:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 229:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 230:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 231:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 232:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 233:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 234:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 235:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 236:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 237:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 238:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 239:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 239 and k < 252:
                w.write_column('K240', N, border)
                w.write('J' + str(k), T[19], cell_format7)
                # w.write_column('F240', N, border)
                if k == 240:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 241:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 242:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 243:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 244:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 245:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 246:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 247:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 248:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 249:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 250:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 251:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 251 and k < 264:
                w.write_column('K252', N, border)
                w.write('J' + str(k), T[20], cell_format8)
                # w.write_column('F252', N, border)
                if k == 252:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 253:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 254:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 255:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 256:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 257:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 258:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 259:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 260:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 261:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 262:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 263:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 263 and k < 276:
                w.write_column('K264', N, border)
                w.write('J' + str(k), T[21], cell_format9)
                # w.write_column('F264', N, border)
                if k == 264:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 265:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 266:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 267:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 268:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 269:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 270:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 271:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 272:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 273:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 274:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 275:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 275 and k < 288:
                w.write_column('K276', N, border)
                w.write('J' + str(k), T[22], cell_format10)
                # w.write_column('F276', N, border)
                if k == 276:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 277:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 278:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 279:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 280:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 281:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 282:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 283:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 284:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 285:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 286:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 287:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 287 and k < 300:
                w.write_column('K288', N, border)
                w.write('J' + str(k), T[23], cell_format11)
                # w.write_column('F288', N, border)
                if k == 288:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 289:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 290:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 291:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 292:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 293:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 294:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 295:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 296:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 297:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 298:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 299:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 299 and k < 312:
                w.write_column('K300', N, border)
                w.write('J' + str(k), T[24], cell_format)
                if k == 300:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 301:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 302:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 303:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 304:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 305:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 306:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 307:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 308:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 309:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 310:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 311:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 311 and k < 324:
                w.write_column('K312', N, border)
                w.write('J' + str(k), T[25], cell_format1)
                if k == 312:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 313:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 314:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 315:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 316:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 317:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 318:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 319:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 320:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 321:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 322:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 323:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 323 and k < 336:
                w.write_column('K324', N, border)
                w.write('J' + str(k), T[26], cell_format2)
                if k == 324:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 325:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 326:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 327:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 328:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 329:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 330:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 331:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 332:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 333:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 334:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 335:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 335 and k < 348:
                w.write_column('K336', N, border)
                w.write('J' + str(k), T[27], cell_format3)
                if k == 336:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 337:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 338:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 339:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 340:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 341:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 342:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 343:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 344:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 345:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 346:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 347:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 347 and k < 360:
                w.write_column('K348', N, border)
                w.write('J' + str(k), T[28], cell_format4)
                if k == 348:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 349:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 350:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 351:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 352:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 353:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 354:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 355:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 356:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 357:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 358:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 359:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 359 and k < 372:
                w.write_column('K360', N, border)
                w.write('J' + str(k), T[29], cell_format5)
                if k == 360:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 361:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 362:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 363:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 364:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 365:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 366:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 367:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 368:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 369:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 370:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 371:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 371 and k < 384:
                w.write_column('K372', N, border)
                w.write('J' + str(k), T[30], cell_format6)
                if k == 372:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 373:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 374:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 375:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 376:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 377:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 378:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 379:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 380:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 381:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 382:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 383:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 383 and k < 396:
                w.write_column('K384', N, border)
                w.write('J' + str(k), T[31], cell_format7)
                if k == 384:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 385:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 386:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 387:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 388:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 389:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 390:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 391:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 392:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 393:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 394:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 395:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 395 and k < 408:
                w.write_column('K396', N, border)
                w.write('J' + str(k), T[32], cell_format8)
                if k == 396:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 397:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 398:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 399:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 400:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 401:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 402:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 403:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 404:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 405:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 406:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 407:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 407 and k < 420:
                w.write_column('K408', N, border)
                w.write('J' + str(k), T[33], cell_format9)
                if k == 408:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 409:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 410:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 411:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 412:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 413:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 414:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 415:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 416:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 417:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 418:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 419:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 419 and k < 432:
                w.write_column('K420', N, border)
                w.write('J' + str(k), T[34], cell_format10)
                if k == 420:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 421:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 422:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 423:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 424:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 425:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 426:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 427:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 428:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 429:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 430:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 431:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 431 and k < 444:
                w.write_column('K432', N, border)
                w.write('J' + str(k), T[35], cell_format11)
                if k == 432:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 433:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 434:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 435:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 436:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 437:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 438:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 439:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 440:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 441:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 442:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 443:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 443 and k < 456:
                w.write_column('K444', N, border)
                w.write('J' + str(k), T[36], cell_format)
                if k == 444:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 445:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 446:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 447:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 448:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 449:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 450:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 451:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 452:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 453:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 454:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 455:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 455 and k < 468:
                w.write_column('K456', N, border)
                w.write('J' + str(k), T[37], cell_format1)
                if k == 456:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 457:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 458:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 459:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 460:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 461:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 462:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 463:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 464:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 465:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 466:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 467:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 467 and k < 480:
                w.write_column('K468', N, border)
                w.write('J' + str(k), T[38], cell_format2)
                if k == 458:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 469:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 470:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 471:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 472:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 473:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 474:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 475:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 476:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 477:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 478:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 479:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 479 and k < 492:
                w.write_column('K480', N, border)
                w.write('J' + str(k), T[39], cell_format3)
                if k == 480:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 481:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 482:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 483:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 484:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 485:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 486:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 487:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 488:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 489:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 490:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 491:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 491 and k < 504:
                w.write_column('K492', N, border)
                w.write('J' + str(k), T[40], cell_format4)
                if k == 492:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 493:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 494:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 495:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 496:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 497:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 498:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 499:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 500:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 501:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 502:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 503:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 503 and k < 516:
                w.write_column('K504', N, border)
                w.write('J' + str(k), T[41], cell_format5)
                if k == 504:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 505:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 506:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 507:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 508:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 509:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 510:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 511:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 512:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 513:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 514:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 515:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 515 and k < 528:
                w.write_column('K516', N, border)
                w.write('J' + str(k), T[42], cell_format6)
                if k == 516:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 517:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 518:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 519:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 520:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 521:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 522:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 523:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 524:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 525:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 526:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 527:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 527 and k < 540:
                w.write_column('K528', N, border)
                w.write('J' + str(k), T[43], cell_format7)
                if k == 528:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 529:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 530:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 531:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 532:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 533:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 534:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 535:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 536:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 537:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 538:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 539:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 539 and k < 552:
                w.write_column('K540', N, border)
                w.write('J' + str(k), T[44], cell_format8)
                if k == 540:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 541:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 542:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 543:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 544:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 545:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 546:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 547:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 548:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 549:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 550:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 551:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 551 and k < 564:
                w.write_column('K552', N, border)
                w.write('J' + str(k), T[45], cell_format9)
                if k == 552:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 553:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 554:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 555:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 556:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 557:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 558:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 559:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 560:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 561:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 562:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 563:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 563 and k < 576:
                w.write_column('K564', N, border)
                w.write('J' + str(k), T[46], cell_format10)
                if k == 564:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 565:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 566:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 567:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 568:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 569:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 570:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 571:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 572:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 573:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 574:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 575:
                    w.write('I' + str(k), N[11], cell_format11)
            elif k > 575 and k < 588:
                w.write_column('K576', N, border)
                w.write('J' + str(k), T[47], cell_format11)
                if k == 576:
                    w.write('I' + str(k), N[0], cell_format)
                elif k == 577:
                    w.write('I' + str(k), N[1], cell_format1)
                elif k == 578:
                    w.write('I' + str(k), N[2], cell_format2)
                elif k == 579:
                    w.write('I' + str(k), N[3], cell_format3)
                elif k == 580:
                    w.write('I' + str(k), N[4], cell_format4)
                elif k == 581:
                    w.write('I' + str(k), N[5], cell_format5)
                elif k == 582:
                    w.write('I' + str(k), N[6], cell_format6)
                elif k == 583:
                    w.write('I' + str(k), N[7], cell_format7)
                elif k == 584:
                    w.write('I' + str(k), N[8], cell_format8)
                elif k == 585:
                    w.write('I' + str(k), N[9], cell_format9)
                elif k == 586:
                    w.write('I' + str(k), N[10], cell_format10)
                elif k == 587:
                    w.write('I' + str(k), N[11], cell_format11)

        k += 1

    j += 1
workbook.close()
'''
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open(r'C:\\Users\\BE24\\Desktop\\Projet Fibre 21\\PDS\\21_011_327.xlsx')
auto = 0
while auto < l :
    ws = wb.Worksheets(c_extremite[auto])
    ws.Columns.AutoFit()
    auto += 1
wb.Save()
'''
# excel.Application.Quit()
