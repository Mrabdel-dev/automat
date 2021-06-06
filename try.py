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
import xlsxwriter
rootBook = xlsxwriter.Workbook('fileGenerated/hh.xlsx')
wr = rootBook.add_worksheet()
wr.write(0,0,"hello")
rootBook.close()