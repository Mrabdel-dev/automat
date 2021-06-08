# def duplicates(lst, item):
#     return [i for i, x in enumerate(lst) if x == item]
#
#
#
# listi = ['1','2','158','99','93']
# for i in enumerate(listi):
#     print(i)
# c = []
# for i in range(65, 77):
#     c.append(chr(i))
# c.append(chr(78))
#
# print(c[15])
# import xlsxwriter
# rootBook = xlsxwriter.Workbook('fileGenerated/hh.xlsx')
# wr = rootBook.add_worksheet()
# wr.write(0,0,"hello")
# rootBook.close()
# import operator
#
# az ={'rrrrr':144,'hhhhhh':568,'yyyyyyyyy':144}
# sortedOne =dict(sorted(az.items(), key=operator.itemgetter(1)))
#
# sh=sortedOne.keys()
# print(sortedOne)
# for sh in sh :
#     print(sh)
#                     wr.write(p, column, state, border)
#                     column = column + 1
#                     # CAS VALUE
#                     x = nextBoiteSheet.cell(row=s, column=7).value
#                     wr.write(p, column, x, cassette)
#                     column = column + 1
#                     if state != 'A STOCKER' and state != 'LIBRE':
#                         # TUBE VALUE
#                         x = nextBoiteSheet.cell(row=s, column=10).value
#                         wr.write(p, column, x, stringCassette(str(x)))
#                         column = column + 1
#                         # FIBRE VALUE
#                         x = nextBoiteSheet.cell(row=s, column=9).value
#                         wr.write(p, column, x, stringCassette(str(x)))
#                         column = column + 1
#                         # CABLE VALUE 2
#                         x = nextBoiteSheet.cell(row=s, column=14).value
#                         wr.write(p, column, x, border)
#                         column = column + 1
#                         # BOITE VALUE 2
#                         x = str(nextBoiteSheet.cell(row=s, column=14).value)
#                         boite = x[-4:]
#                         if boite is not None:
#                             boit = getBoiteName(boite)
#                             wr.write(p, column, boit, border)
#                             column = column + 1
#                             newBoite = str(boit)
#                     else:
#                         done = False