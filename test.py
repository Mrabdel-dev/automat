import csv

# ##################################### const declaration ##############
import operator

import xlsxwriter

cableName = []
tubeNumberI = []
fibreNumberI = []
boiteName = []
casseteName = []
destinationCable = []
tubeNumberII = []
fibreNumberII = []
etat = []
SroSet = set()
sroCable = []
cap = 0
found = True
x = 0
rowmax = 0
# ##########################################################################################
with open('fileGenerated/21_017_101_EPISSURAGE_.csv', 'rt')as f:
    data = csv.DictReader(f, delimiter=';')
    # #<---------- get the value from csv epesourge table---------------------------------->#
    for row in data:
        cableName.append(row['ï»¿CODE_CABLE_ORIGINE'])
        tubeNumberI.append(row['NUMERO_TUBE_ORIGINE'])
        fibreNumberI.append(row['NUMERO_FIBRE_ORIGINE'])
        boiteName.append(row['CODE_BOITE'])
        casseteName.append(row['CODE_CASSETTE'])
        destinationCable.append(row['CODE_CABLE_DESTINATION'])
        tubeNumberII.append(row['NUMERO_TUBE_DESTINATION'])
        fibreNumberII.append(row['NUMERO_FIBRE_DESTINATION'])
        etat.append(row['ETAT'])
        rowmax += 1
    # ##########################################################################################
    # get sro cable ande there capcity
    for i in cableName:
        for j in destinationCable:
            if i == j:
                found = True
                break
            else:
                found = False
        if not found:
            SroSet.add(i)

    dictCable = {}
    for cb in SroSet:
        cap = 0
        for j in cableName:
            if cb == j:
                cap += 1
            else:
                continue
        dictCable.update({cb: cap})

sortedSro = dict(sorted(dictCable.items(), key=operator.itemgetter(1)))
sroCable = list(sortedSro.keys())
print(dictCable)
print(sroCable)
# ########################################################################################
# <-----------------------route file creation------------------------------------------->
rootBook = xlsxwriter.Workbook('Root-101.xlsx')
wr = rootBook.add_worksheet()
# define the character and style of cell inside excel
bold = rootBook.add_format({'bold': True, "border": 1})
bold1 = rootBook.add_format({'bold': True})
border = rootBook.add_format({"border": 1})
header = rootBook.add_format({'bold': True, 'border': 1, 'bg_color': '#037d50'})
cassette = rootBook.add_format({"bg_color": '#A9A9A9', "border": 1})
cell_formatCapacity = rootBook.add_format({"bg_color": '#E6E6FA', "border": 1})
cell_format1 = rootBook.add_format({"bg_color": 'red', "border": 1})
cell_format2 = rootBook.add_format({"bg_color": 'blue', "border": 1})
cell_format3 = rootBook.add_format({"bg_color": '#00FF00', "border": 1})
cell_format4 = rootBook.add_format({"bg_color": 'yellow', "border": 1})
cell_format5 = rootBook.add_format({"bg_color": '#BF00FF', "border": 1})
cell_format6 = rootBook.add_format({"bg_color": 'white', "border": 1})
cell_format7 = rootBook.add_format({"bg_color": '#FFBF00', "border": 1})
cell_format8 = rootBook.add_format({"bg_color": '#828282', "border": 1})
cell_format9 = rootBook.add_format({"bg_color": '#816B56', "border": 1})
cell_format10 = rootBook.add_format({"bg_color": '#333333', "border": 1})
cell_format11 = rootBook.add_format({"bg_color": '#00FFBF', "border": 1})
cell_format12 = rootBook.add_format({"bg_color": '#FFAAD4', "border": 1})
colorList = [cell_format1, cell_format2, cell_format3, cell_format4, cell_format5, cell_format6, cell_format7,
             cell_format8, cell_format9, cell_format10, cell_format11, cell_format12, border]


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


def baseHeader():
    wr.write('A1', 'SRO', header)
    wr.write('B1', 'P', header)
    wr.write('C1', 'C ', header)
    wr.write('D1', 'L', header)
    wr.write('E1', 'TIROIR', header)
    wr.write('F1', 'TYPE', header)


def normalHeader(j):
    for i in range(j, 20):
        wr.write(0, j, 'CAS', header)
        j = j + 1
        wr.write(0, j, 'T', header)
        j = j + 1
        wr.write(0, j, 'F', header)
        j = j + 1
        wr.write(0, j, 'CABLE', header)
        j = j + 1
        wr.write(0, j, 'BOITE', header)
        j = j + 1
        wr.write(0, j, 'TYPE', header)
        j = j + 1


def getIndex(val, tabl: list):
    for i in tabl:
        if i == val:
            return tabl.index(val)


def getNextIndex(val, tube, fibre, tabl: list):
    for i in range(0, rowmax):
        c = tabl[i]
        t = tubeNumberI[i]
        f = fibreNumberI[i]
        if c == val and tube == t and fibre == f:
            return i


OK = True


def checkPassage(index):
    global OK
    tube = tubeNumberII[index]
    fibre = fibreNumberII[index]
    try:
        while OK:
            boite = boiteName[index]
            ETAT = etat[index]
            cassete = casseteName[index]
            cable2 = destinationCable[index]
            if cable2 is not None:
                if ETAT == 'STOCKEE':
                    if cassete == 'FOND DE BOITE' or boite.startswith('PEC'):
                        return 'LIB'
                    else:
                        return 'PASS'
                elif ETAT == 'EPISSUREE':
                    return 'PASS'
                else:
                    # print(ETAT)
                    # print(tube, fibre, cable2)
                    index = getNextIndex(cable2, tube, fibre, cableName)
                    # print(index)
            else:
                return 'LIB'
    except TypeError:
        print(index, 'eroooooooooor')


#
# x = checkPassage(100)
# y = getNextIndex('CDI-21-017-101-3007', 3, 4, cableName)
# print(y)
# # print(x)
# print(destinationCable.index('CDI-21-017-101-3005'))
# print(cableName[4704])

# getNextIndex('CDI-21-017-101-1004',9,4,cableName)#######################################################################################


# <----------------------------------generation part from epesourage file to route ------------------>
# all constant we work with inside
p = 0
c = []
for i in range(65, 77):
    c.append(chr(i))
c.append(chr(78))
L = 1
T = 0
f = 0
N = 0
Line = 0
done = True
b = 0
Lin = 1
l = 1
column = 0
TEST = ''
SRO = 'SRO' + str(cableName[0])[4:14]
# # ##################################
baseHeader()
normalHeader(6)
for cab in sroCable:
    T += 1
    Line = getIndex(cab, cableName)
    done = True
    while done:
        TEST = checkPassage(Line)
        print('#' * 25)
        print(TEST)
        print(cableName[Line])
        if TEST == 'PASS':
            if f == 13:
                f = 0
            b = l + 1
            wr.write('A' + str(b), SRO, border)
            wr.write('B' + str(b), l, border)
            wr.write('E' + str(b), 'TIROIR_' + str(T), border)
            wr.write('F' + str(b), 'CONNECTEUR', border)
            wr.write('C' + str(b), c[f], border)
            wr.write('D' + str(b), L, border)
            f = f + 1
            if l % 24 == 0:
                if L == 6:
                    L = 1
                elif L < 6:
                    L = L + 1
            l += 1
            # #################get the values#############################
            CAS = tubeNumberI[Line]
            tube1 = tubeNumberI[Line]
            fibre1 = fibreNumberI[Line]
            cable1 = cableName[Line]
            boite = boiteName[Line]
            ETAT = etat[Line]
            cassete = casseteName[Line]
            tube2 = tubeNumberII[Line]
            fibr2 = fibreNumberII[Line]
            cable2 = destinationCable[Line]
            column = 6
            # CAS VALUE
            wr.write(Lin, column, CAS, cassette)
            column = column + 1
            # TUBE1 VALUE
            wr.write(Lin, column, tube1, stringCassette(str(tube1)))
            column = column + 1
            # FIBRE1 VALUE
            wr.write(Lin, column, fibre1, stringCassette(str(fibre1)))
            column = column + 1
            # CABLE1 VALUE
            wr.write(Lin, column, cable1, border)
            column = column + 1
            # BOITE1 VALUE
            wr.write(Lin, column, boite, border)
            column = column + 1
            # TYPE VALUE
            wr.write(Lin, column, ETAT, border)
            column = column + 1
            # Cassete VALUE
            wr.write(Lin, column, cassete, cassette)
            column = column + 1
            # TUBE2 VALUE
            wr.write(Lin, column, tube2, stringCassette(str(tube2)))
            column = column + 1
            # FIBRE2 VALUE
            wr.write(Lin, column, fibr2, stringCassette(str(fibr2)))
            column = column + 1
            # CABLE2 VALUE 2
            wr.write(Lin, column, cable2, border)
            column = column + 1
            if ETAT == 'STOCKEE':
                # y = getNextIndex(cable2, tube2, fibr2, cableName)
                # boite = boiteName[y]
                # # BOITE2 VALUE
                # wr.write(Lin, column, boite, border)
                # column = column + 1
                # ETAT = etat[y]
                # # TYPE2 VALUE
                # wr.write(Lin, column, ETAT, border)
                # column = column + 1
                # cassete = casseteName[y]
                # # Cassete2 VALUE
                # wr.write(Lin, column, cassete, cassette)
                # column = column + 1
                keep = False
                Line += 1
            else:
                keep = True
                while keep:
                    try:
                        # print(cable2, tube2, fibr2)
                        y = getNextIndex(cable2, tube2, fibr2, cableName)
                        boite = boiteName[y]
                        # BOITE2 VALUE
                        wr.write(Lin, column, boite, border)
                        column = column + 1
                        ETAT = etat[y]
                        # TYPE2 VALUE
                        wr.write(Lin, column, ETAT, border)
                        column = column + 1
                        cassete = casseteName[y]
                        # Cassete2 VALUE
                        wr.write(Lin, column, cassete, cassette)
                        column = column + 1
                        # TUBE2 VALUE
                        tube2 = tubeNumberII[y]
                        wr.write(Lin, column, tube2, stringCassette(str(tube2)))
                        column = column + 1
                        # FIBRE2 VALUE
                        fibr2 = fibreNumberII[y]
                        wr.write(Lin, column, fibr2, stringCassette(str(fibr2)))
                        column = column + 1
                        # CABLE2 VALUE 2
                        cable2 = destinationCable[y]
                        wr.write(Lin, column, cable2, border)
                        column = column + 1
                    except TypeError:
                        print(cable2, tube2, fibr2)
                        continue
                    if ETAT == 'STOCKEE':
                        # y = getNextIndex(cable2, tube2, fibr2, cableName)
                        # boite = boiteName[y]
                        # # BOITE2 VALUE
                        # wr.write(Lin, column, boite, border)
                        # column = column + 1
                        # ETAT = etat[y]
                        # # TYPE2 VALUE
                        # wr.write(Lin, column, ETAT, border)
                        # column = column + 1
                        # cassete = casseteName[y]
                        # # Cassete2 VALUE
                        # wr.write(Lin, column, cassete, cassette)
                        # column = column + 1
                        keep = False
                        Line += 1
                    else:
                        continue
            Lin += 1
        else:
            Line += 1
        cabel = cableName[Line]
        if cab != cabel:
            done = False
rootBook.close()
