from openpyxl import load_workbook
# load the old book you gona modify
workbookOld = load_workbook('fileGenerated/Export projets Free 04-05.xlsx')
oldBook = workbookOld.active
# load the new book that you want get value from it
workbookNew = load_workbook('fileGenerated/export_projets_2021 IDF ( Nouveau).xlsx')
newBook = workbookNew.active
# define parameter for loop
maxRow = newBook.max_row
print('######## the number of values in the new file########')
print(maxRow-1)
maxCol = newBook.max_column
maxRowOld = oldBook.max_row
print('######## the number of values in the old file ########')
print(maxRowOld-1)
maxColOld = oldBook.max_column
NbrAdd = 0
NbrExs = 0
NbrDel = 0
# the first loop is add all the new value that doesn't exist in the old file
for i in range(2, maxRow+1):
    valNew = newBook.cell(row=i, column=4).value
    for j in range(2, maxRowOld+1):
        valOld = oldBook.cell(row=j, column=4).value
        if valNew == valOld:
            found = 1
            NbrExs = NbrExs + 1
            break
        else:
            found = 0
    if found == 0:
        modRowOld = oldBook.max_row
        u = modRowOld+1
        NbrAdd = NbrAdd+1
        for k in range(1, 10):
            valN = newBook.cell(row=i, column=k).value
            oldBook.cell(row=u, column=k).value = valN
newMax = oldBook.max_row
# the second loop is for delete all value doesn't come with new file
for c in range(2, newMax):
    valTest = oldBook.cell(row=c, column=4).value
    for d in range(2, maxRow+1):
        valTestedWith = newBook.cell(row=d, column=4).value
        if valTest == valTestedWith:
            y = 1
            break
        else:
            y = 0
    if y == 0:
        NbrDel = NbrDel+1
        for e in range(1, maxColOld+1):
            oldBook.cell(row=c, column=e).value = ''
print('##### statistique of operation ##########')
print(f'number of element added is : ==>{NbrAdd}')
print(f'number of element deleted is : ==>{NbrDel}')
print(f'number of element exist already is : ==>{NbrExs}')
print('#################### the number of values in the old file after update ####################')
print(oldBook.max_row -1)
# save the file
fileN = 'fileGenerated/NewUpDl1.xlsx'
workbookOld.save(fileN)