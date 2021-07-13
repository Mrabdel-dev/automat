from openpyxl import load_workbook

# load the C6 book
workbook = load_workbook('F99999JJMMAA_C6 (1).xlsx')
BookC6 = workbook['Export 2']
# load the C3A book
workbook1 = load_workbook('C3A.xlsx')
BookC3A = workbook1['Commandes Fermes']
###################################################################################
codeInsee = str(BookC6.cell(row=3, column=7).value)
NappuiOrigi = []
typeApp = []
typeEL = []
NomCab = []
longeur = []
NappuiDist = []
nature = []
maxRow = BookC6.max_row
x = 0
y = 0
for i in range(9, maxRow):
    N = BookC6.cell(row=i, column=1).value
    T = BookC6.cell(row=i, column=2).value
    if N is not None and T is not None:
        x = N
        y = str(T)
        NappuiOrigi.append(x)
        typeApp.append(y)
    else:
        NappuiOrigi.append(x)
        typeApp.append(y)
    if y == 'ORT' or y == 'EDF':
        typeEL.append('AT')
    else:
        typeEL.append('A')
    NomCab.append(str(BookC6.cell(row=i, column=19).value))
    longeur.append(BookC6.cell(row=i, column=20).value)
    NappuiDist.append(BookC6.cell(row=i, column=25).value)
    nature.append(BookC6.cell(row=i, column=32).value)

def getDiameter(nomcab):
    return nomcab
def checkFill(NomDist,index):
    test =True
    for i in range(0,index):
        if NomDist == NappuiOrigi[i]:
            test = False
            break
    return test
Len = len(NomCab)
Lin = 15
for i in range(0,Len):
    numC = NomCab[i]
    if numC.startswith('A'):
        orig = str(NappuiOrigi[i])+"/"+codeInsee
        test = checkFill(NappuiDist[i],i)
        dist = str(NappuiDist[i])+"/"+codeInsee
        if test:
            BookC3A.cell(row=Lin,column=2).value=typeEL[i]
            BookC3A.cell(row=Lin, column=3).value =orig
            BookC3A.cell(row=Lin, column=4).value =typeEL[i]
            BookC3A.cell(row=Lin, column=5).value =dist
            BookC3A.cell(row=Lin, column=5).value = longeur[i]


