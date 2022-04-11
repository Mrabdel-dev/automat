import xlsxwriter
from dbfread import DBF

workbook = xlsxwriter.Workbook('dbl/dblcompareNew.xlsx')
bold = workbook.add_format({'bold': True, "border": 1})
w = workbook.add_worksheet('result')
border = workbook.add_format({"border": 1})
w.write("A1", "refOLD", bold)
w.write("B1", "nbrprise", bold)
w.write("C1", "techno", bold)
w.write("D1", "typecli", bold)
w.write("F1", "refOLD", bold)
w.write("G1", "nbrprise", bold)
w.write("H1", "techno", bold)
w.write("I1", "typecli", bold)
zaPboDbl = DBF('dbl/21_011_069_EXE_CCE_DBL_MAN_001_B_AI.dbf', load=True, encoding='iso-8859-1')
zapLen = len(zaPboDbl)
# FROM THE JOIN ZAPBO AND DBL
ref1 = []
nbPrise1 = []
tECHNO1 = []
typeBat1 = []
typeCli1 = []
statut1 = []
for k in range(0, zapLen):
    ref1.append(str(zaPboDbl.records[k]['REF_IMB']))
    nbPrise1.append(str(zaPboDbl.records[k]['NB_PRISE']))
    tECHNO1.append(str(zaPboDbl.records[k]['TECHNO']))
    typeBat1.append(str(zaPboDbl.records[k]['TYPE_BAT']))
    typeCli1.append(str(zaPboDbl.records[k]['TYPE_CLI']))
    statut1.append(str(zaPboDbl.records[k]['STATUT']))



zaPboDbl2 = DBF('dbl/21_011_069_EXE_CCE_DBL_MAN_001_B.dbf', load=True, encoding='iso-8859-1')
zapLen2 = len(zaPboDbl2)
# FROM THE JOIN ZAPBO AND DBL
ref2 = []
nbPrise2 = []
tECHNO2 = []
typeBat2 = []
typeCli2 = []
statut2 = []
for k in range(0, zapLen2):
    ref2.append(str(zaPboDbl2.records[k]['REF_IMB']))

    nbPrise2.append(str(zaPboDbl2.records[k]['NB_PRISE']))
    tECHNO2.append(str(zaPboDbl2.records[k]['TECHNO']))
    typeBat2.append(str(zaPboDbl2.records[k]['TYPE_BAT']))
    typeCli2.append(str(zaPboDbl2.records[k]['TYPE_CLI']))
    statut2.append(str(zaPboDbl2.records[k]['STATUT']))

i = 2
k = 0
for re in ref1:
    k += 1
    index = ref1.index(re)
    nb1 = str(nbPrise1[index])
    tec1 = str(tECHNO1[index])
    typc1 = str(typeCli1[index])
    if str(re).startswith("N"):
        continue
    try:
        ind = ref2.index(re)
        ref = ref2[ind]
        nb2 = str(nbPrise2[ind])
        tec2 = str(tECHNO2[ind])
        typc2 = str(typeCli2[ind])
        if nb1 != nb2 or tec1 != tec2 or typc1 != typc2:
            print(ref, " ", nb1, " ", ref, " ", nb2)
            w.write("A" + str(i), re, border)
            w.write("B" + str(i), nb1, border)
            w.write("C" + str(i), tec1, border)
            w.write("D" + str(i), typc1, border)
            w.write("F" + str(i), ref, border)
            w.write("G" + str(i), nb2, border)
            w.write("H" + str(i), tec2, border)
            w.write("I" + str(i), typc2, border)
            i += 1

    except :

        w.write("A" + str(i), re, border)
        w.write("B" + str(i), nb1, border)
        w.write("C" + str(i), tec1, border)
        w.write("D" + str(i), typc1, border)
        i += 1


workbook.close()
print(i,k, len(ref1), zapLen, zapLen2)

