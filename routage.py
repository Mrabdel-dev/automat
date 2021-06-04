import xlsxwriter
from openpyxl import load_workbook

rootBook = xlsxwriter.Workbook('fileGenerated/roote.xlsx')
wr = rootBook.add_worksheet()
# define the character and style of cell inside excel
bold = rootBook.add_format({'bold': True, "border": 1})
bold1 = rootBook.add_format({'bold': True})
border = rootBook.add_format({"border": 1})
header = rootBook.add_format({'bold': True, 'border': 1, 'bg_color': '#037d50'})
back = rootBook.add_format({"bg_color": '#CD5C5C', "border": 1})
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
             cell_format8, cell_format9, cell_format10, cell_format11, cell_format12]


def baseHeader():
    wr.write('A1', 'SRO', header)
    wr.write('B1', 'P', header)
    wr.write('C1', 'C ', header)
    wr.write('D1', 'L', header)
    wr.write('E1', 'TIROIR', header)
    wr.write('F1', 'TYPE', header)

def normalHeader(i):
    wr.write('G'+str(i), 'CAS', header)
    wr.write('H'+str(i), 'T', header)
    wr.write('I'+str(i), 'F', header)
    wr.write('J'+str(i), 'CABLE', header)
    wr.write('K'+str(i), 'BOITE', header)
    wr.write('L'+str(i), 'TYPE', header)
pdsBook = load_workbook('fileGenerated/PDS.xlsx')
pdsSheets = pdsBook.sheetnames
sheetSro = []
cableSro = []
capSro= []
value = ''
for sh in pdsSheets:
    sheet = pdsBook[sh]
    print(sh)
    print(sheet.max_row-11)
    print('#'*15)
    value = sheet.cell(row=1, column=1).value
    if str(value).startswith('SRO'):
        sheetSro.append(sh)

        SRO = value
        cableSro.append(str(sheet.cell(row=12, column=1).value))
        capSro.append(int(sheet.cell(row=12, column=3).value))

    else:
        pass
baseHeader()
normalHeader(1)
p=0
c=[]
for i in range(65,77):
    c.append(chr(i))
c.append(chr(78))
L=0
T=0
f=0
N=2
Len=0
for b in sheetSro:
    bshet = pdsBook[b]
    L=1
    T=T+1
    Len = Len + bshet.max_row - 11
    for p in range(N,Len+2):
        wr.write('A'+str(p),value, border)
        wr.write('B' + str(p),p, border)
        wr.write('E' + str(p), 'TIROIR_'+str(T), border)
        wr.write('F' + str(p), 'CONNECTEUR', border)
        wr.write('C' + str(p), c[f], border)
        wr.write('D' + str(p), L, border)

        if p%6 ==0:
            if L==6:
                L=1
            elif L <6:
                L = L+1
        else :
            pass



